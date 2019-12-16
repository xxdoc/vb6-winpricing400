VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReArrangeDoc 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmRearrangeDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   6405
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   11298
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1320
         Width           =   1875
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   2
         Top             =   4590
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   8
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
         TabIndex        =   3
         Top             =   4920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   2655
         Left            =   1860
         TabIndex        =   13
         Top             =   1920
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   4683
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRearrangeDoc.frx":27A2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin prjFarmManagement.uctlTextBox txtYear 
         Height          =   435
         Left            =   3720
         TabIndex        =   1
         Top             =   1320
         Width           =   975
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin Threed.SSCheck chkCheckBalance 
         Height          =   435
         Left            =   7080
         TabIndex        =   15
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblPartGroup 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   14
         Top             =   1980
         Width           =   1575
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1890
         TabIndex        =   4
         Top             =   5580
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmRearrangeDoc.frx":307C
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   12
         Top             =   5040
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   11
         Top             =   4650
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   5070
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   6
         Top             =   5580
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   6885
         TabIndex        =   5
         Top             =   5580
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmRearrangeDoc.frx":3396
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmReArrangeDoc"
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

Private m_Balances As Collection
Private m_PartItemsDateLocations As Collection
Private m_PartGroups As Collection
Private Sub cboMonth_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
'   Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
   OKClick = True
   Unload Me
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_Employee.EMP_ID = id
      m_Employee.QueryFlag = 1
      If Not glbDaily.QueryEmployee(m_Employee, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Employee.PopulateFromRS(1, m_Rs)
      
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Employee.EMP_ID = id
   m_Employee.AddEditMode = ShowMode
   m_Employee.PASS_STATUS = "Y"
   
   m_Employee.EmpName.AddEditMode = ShowMode
   m_Employee.EName.AddEditMode = ShowMode
      
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditEmployee(m_Employee, IsOK, True, glbErrorLog) Then
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

Private Function GetNextTransaction(Rs1 As ADODB.Recordset, Rs2 As ADODB.Recordset, II As CLotItem, Ei As CLotItem) As String
Dim EofFlag1 As Boolean
Dim EofFlag2 As Boolean
   
   'Export
   EofFlag1 = Rs1.EOF
   If Not Rs1.EOF Then
      Call Ei.PopulateFromRS(1, Rs1)
   End If
   
   'Import
   EofFlag2 = Rs2.EOF
   If Not Rs2.EOF Then
      Call II.PopulateFromRS(1, Rs2)
   End If
      
   If (EofFlag1 And EofFlag2) Then
      GetNextTransaction = ""
   ElseIf (EofFlag1 And (Not EofFlag2)) Then
      GetNextTransaction = "I"
      Rs2.MoveNext
   ElseIf ((Not EofFlag1) And EofFlag2) Then
      GetNextTransaction = "E"
      Rs1.MoveNext
   Else
      '===
      'การเรียงลำดับมีผลอย่างมาก
      If Ei.DOCUMENT_DATE = DateSerial(2006, 10, 9) Then
         '''Debug.Print ("")
      End If
      If II.DOCUMENT_DATE = DateSerial(2006, 10, 9) Then
         '''Debug.Print ("")
      End If
      If DateToStringInt(Ei.DOCUMENT_DATE) = DateToStringInt(II.DOCUMENT_DATE) Then
         If Ei.PRIORITY1 = II.PRIORITY1 Then
            If Ei.DOCUMENT_NO = II.DOCUMENT_NO Then
               If Ei.TRANSACTION_SEQ < II.TRANSACTION_SEQ Then
                  GetNextTransaction = "E"
               Else
                  GetNextTransaction = "I"
               End If
            ElseIf Ei.DOCUMENT_NO < II.DOCUMENT_NO Then
               GetNextTransaction = "E"
            Else
               GetNextTransaction = "I"
            End If
         ElseIf Ei.PRIORITY1 < II.PRIORITY1 Then
            GetNextTransaction = "E"
         Else
            GetNextTransaction = "I"
         End If
      ElseIf DateToStringInt(Ei.DOCUMENT_DATE) < DateToStringInt(II.DOCUMENT_DATE) Then
         GetNextTransaction = "E"
      Else
         GetNextTransaction = "I"
      End If 'Document date
      '===
      If GetNextTransaction = "I" Then
         Rs2.MoveNext
      ElseIf GetNextTransaction = "E" Then
         Rs1.MoveNext
      End If
   End If 'Eof flag
End Function

'Public Function GetBalanceAmount(PartItemID As Long, LocationID As Long, TxSeq As Long, DocDate As Date) As Object
'Dim EI As CExportItem
'Dim II As CImportItem
'Dim TempRs As ADODB.Recordset
'Dim iCount As Long
'
'   Set TempRs = New ADODB.Recordset
'
'   Set EI = New CExportItem
'   Set II = New CImportItem
'
'   EI.EXPORT_ITEM_ID = -1
'   EI.PIG_FLAG = "N"
'   EI.PART_ITEM_ID = PartItemID
'   EI.LOCATION_ID = LocationID
'   EI.FROM_TX_SEQ = -1
'   EI.TO_TX_SEQ = TxSeq
'   EI.FROM_DATE = -1
'   EI.TO_DATE = DocDate
'   EI.OrderBy = 11
'   EI.OrderType = 2
'   Call EI.QueryData(1, TempRs, iCount)
'   If Not TempRs.EOF Then
'      Call EI.PopulateFromRS(1, TempRs)
'   End If
'
'   II.IMPORT_ITEM_ID = -1
'   II.PIG_FLAG = "N"
'   II.PART_ITEM_ID = PartItemID
'   II.LOCATION_ID = LocationID
'   II.FROM_TX_SEQ = -1
'   II.TO_TX_SEQ = TxSeq
'   II.FROM_DATE = -1
'   II.TO_DATE = DocDate
'   II.OrderBy = 12
'   II.OrderType = 2
'   Call II.QueryData(1, TempRs, iCount)
'   If Not TempRs.EOF Then
'      Call II.PopulateFromRS(1, TempRs)
'   End If
'
'   If EI.TRANSACTION_SEQ > II.TRANSACTION_SEQ Then
'      Set GetBalanceAmount = EI
'   Else
'      Set GetBalanceAmount = II
'   End If
'
'   If TempRs.State = adStateOpen Then
'      Call TempRs.Close
'   End If
'   Set TempRs = Nothing
'   Set EI = Nothing
'   Set II = Nothing
'End Function

Private Sub GetRelateItem1(II As CLotItem, Ei As CLotItem)
Dim iCount As Long
Dim TempRs As ADODB.Recordset

   Set TempRs = New ADODB.Recordset
   
   Ei.LOT_ITEM_ID = -1
   Ei.GUI_ID = II.GUI_ID
   Ei.TX_TYPE = "E"
   Call Ei.QueryData(1, TempRs, iCount, False)
   If Not TempRs.EOF Then
      Call Ei.PopulateFromRS(1, TempRs)
   End If
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub

Private Sub GeneratePartItemLocationDate(O As Object, ImpI As CLotItem)
Dim Key As String
Dim Ba As CBalanceAccum
Dim II As CLotItem
Dim TempII As CLotItem
Dim AvgPrice As Double
   
   Key = O.PART_ITEM_ID & "-" & O.LOCATION_ID & "-" & DateToStringInt(O.DOCUMENT_DATE)
   Set II = GetImportItem(m_PartItemsDateLocations, Key)
   If II.PART_ITEM_ID <= 0 Then
      Set TempII = New CLotItem
      TempII.LOCATION_ID = O.LOCATION_ID
      TempII.PART_ITEM_ID = O.PART_ITEM_ID
      TempII.DOCUMENT_DATE = O.DOCUMENT_DATE
      TempII.BALANCE_AMOUNT = ImpI.CURRENT_AMOUNT
      TempII.TOTAL_INCLUDE_PRICE = ImpI.TOTAL_INCLUDE_PRICE
      TempII.INCLUDE_UNIT_PRICE = MyDiffEx(ImpI.TOTAL_INCLUDE_PRICE, ImpI.CURRENT_AMOUNT)
      If O.TX_TYPE = "I" Then
         TempII.ALL_IMPORT_AMT = O.IMPORT_AMOUNT
      ElseIf O.TX_TYPE = "E" Then
         TempII.ALL_EXPORT_AMT = O.EXPORT_AMOUNT
      End If
      Call m_PartItemsDateLocations.add(TempII, Key)
      Set TempII = Nothing
   Else
      If O.TX_TYPE = "I" Then
         II.ALL_IMPORT_AMT = II.ALL_IMPORT_AMT + O.IMPORT_AMOUNT
         II.BALANCE_AMOUNT = II.BALANCE_AMOUNT + O.IMPORT_AMOUNT
         II.TOTAL_INCLUDE_PRICE = II.TOTAL_INCLUDE_PRICE + O.TOTAL_INCLUDE_PRICE
         'II.INCLUDE_UNIT_PRICE = MyDiffEx(ImpI.TOTAL_INCLUDE_PRICE + O.TOTAL_INCLUDE_PRICE, ImpI.CURRENT_AMOUNT + O.IMPORT_AMOUNT)
         II.INCLUDE_UNIT_PRICE = MyDiffEx(II.TOTAL_INCLUDE_PRICE, II.BALANCE_AMOUNT)
      ElseIf O.TX_TYPE = "E" Then
         II.ALL_EXPORT_AMT = II.ALL_EXPORT_AMT + O.EXPORT_AMOUNT
         II.BALANCE_AMOUNT = II.BALANCE_AMOUNT - O.EXPORT_AMOUNT
         If O.ADJUST_FLAG = "Y" Then
            II.TOTAL_INCLUDE_PRICE = II.TOTAL_INCLUDE_PRICE - O.TOTAL_INCLUDE_PRICE ' (O.EXPORT_AMOUNT * ImpI.INCLUDE_UNIT_PRICE)
            II.INCLUDE_UNIT_PRICE = O.NEED_AVG_PRICE 'O.INCLUDE_UNIT_PRICE
         Else
            II.TOTAL_INCLUDE_PRICE = II.TOTAL_INCLUDE_PRICE - (O.EXPORT_AMOUNT * ImpI.INCLUDE_UNIT_PRICE)
         End If
      End If
   End If
End Sub

'Public Function GetBalanceItem(Col As Collection, PartItemID As Long, LocationID As Long, DocDate As Date) As Object
'Dim D As Object
'Dim Key As String
'Dim MaxSeq As Long
'Dim i As Long
'Dim MaxIndex As Long
'Static II As CImportItem
'Dim MaxDate As Date
'
'   MaxDate = -2
'   For Each D In Col
''''Debug.Print D.TX_TYPE & ";" & D.PART_ITEM_ID & ";" & D.LOCATION_ID & ";" & DateToStringInt(D.DOCUMENT_DATE) & ";" & D.CURRENT_AMOUNT
'      If (DateToStringInt(D.DOCUMENT_DATE) < DateToStringInt(DocDate)) And (D.PART_ITEM_ID = PartItemID) And (D.LOCATION_ID = LocationID) Then
'         If DateToStringInt(D.DOCUMENT_DATE) > DateToStringInt(MaxDate) Then
'            MaxDate = InternalDateToDate(DateToStringInt(D.DOCUMENT_DATE))
'         End If
'      End If
'   Next D
'
''If MaxDate <= 0 Then
'''Debug.Print
''End If
'
'   i = 0
'   MaxSeq = -1
'   MaxIndex = -1
'   For Each D In Col
'      i = i + 1
'
'      If (D.PART_ITEM_ID = PartItemID) And (D.LOCATION_ID = LocationID) And _
'         (DateToStringInt(D.DOCUMENT_DATE) = DateToStringInt(MaxDate)) Then
'            If D.TRANSACTION_SEQ > MaxSeq Then
'               MaxSeq = D.TRANSACTION_SEQ
'               MaxIndex = i
'            End If
'      End If
'   Next D
'
'   If MaxIndex > 0 Then
'      Set GetBalanceItem = Col(MaxIndex)
'   Else
'      If II Is Nothing Then
'         Set II = New CImportItem
'      End If
'      Set GetBalanceItem = II
'   End If
'End Function

Private Sub CalculateRMPrice(IvID As Long, Li As CLotItem, InventoryBals As Collection)
Dim Jb As CJob
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Ji As CJobInput
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim Jp As CJobParameter
Dim TempKey As String
Dim TempII As CLotItem

   Set Jb = New CJob
   Set TempRs = New ADODB.Recordset
   
   If IvID = 319559 Then
      'Debug.Print
   End If
   
   Jb.JOB_ID = -1
   Jb.INVENTORY_DOC_ID = IvID
   Jb.QueryFlag = 1
   Call glbProduction.QueryJob(Jb, TempRs, iCount, IsOK, glbErrorLog)
   If Not TempRs.EOF Then
      Call Jb.PopulateFromRS(1, TempRs)
   
      Sum1 = 0
      Sum3 = 0
      For Each Ji In Jb.Inputs
         TempKey = Ji.LOCATION_ID & "-" & Ji.PART_ITEM_ID
         Set TempII = GetImportItem(InventoryBals, TempKey)
         
         If Ji.PARAM_ID <= 0 Then
            '  คิดเป็นต้นทุนวัตถุดิบ
'            Sum1 = Sum1 + TempII.INCLUDE_UNIT_PRICE * Ji.TX_AMOUNT
            Sum1 = Sum1 + Ji.INCLUDE_UNIT_PRICE * Ji.TX_AMOUNT 'ต้องระวังเรื่องลำดับ ต้องให้ E มีลำดับคิวรี่มาก่อน I ที่ DOCUMENT_PRIORITY
         Else
            Sum3 = Sum3 + TempII.INCLUDE_UNIT_PRICE * Ji.TX_AMOUNT
         End If
      Next Ji
      
      Sum2 = 0
      For Each Jp In Jb.Parameters
         Sum2 = Sum2 + Jp.PARAM_AMOUNT
      Next Jp
   End If
   
   Set TempRs = Nothing
   Set Jb = Nothing
   
   Li.INCLUDE_UNIT_PRICE = MyDiffEx(Sum1 + Sum2 + Sum3, Li.IMPORT_AMOUNT)
   Li.TOTAL_INCLUDE_PRICE = Sum1 + Sum2 + Sum3
   Li.RAW_COST = Sum1
   Li.EXPENSE_COST = Sum2 + Sum3
End Sub

Private Function IsSelected(Li As CLotItem) As Boolean
Dim p As CPartGroup

   Set p = GetPartGroup(m_PartGroups, Trim(str(Li.PART_GROUP_ID)))
   IsSelected = (p.SELECT_FLAG = "Y")
End Function

'Private Sub cmdStartOld_Click()
''On Error Resume Next
'Dim Percent As Double
'Dim MIN As Double
'Dim MAX As Double
'Dim RecordCount As Double
'Dim O As Object
'Dim TempO As Object
'Dim InventoryBals As Collection
'Dim RName As String
'Dim cData As CPartLocation
'Dim I As Long
'Dim j As Long
'Dim strFormat As String
'Dim IsOK As Boolean
'Dim Amt As Double
'Dim Ei As CLotItem
'Dim II As CLotItem
'Dim Rs1 As ADODB.Recordset
'Dim Rs2 As ADODB.Recordset
'Dim TxCode As String
'Dim iCount As Long
'Dim AvgPrice As Double
'Dim PrevAmount As Double
'Dim CurrentAmount As Double
'Dim HasBegin As Boolean
'Dim TempII As CLotItem
'Dim TempKey As String
'Dim Count1 As Long
'Dim Count2 As Long
'Dim TempCol As Collection
'Dim TempEi As CLotItem
'Dim ExportTotalPrice As Double
'Dim NewDate As Date
'Dim Ba As CBalanceAccum
'Dim BalanceAccums As Collection
'Dim NewTotalPrice As Double
'Dim IsSelectd As Boolean
'
'   RName = "genDoc"
''-----------------------------------------------------------------------------------------------------
''                                             Query Here
''-----------------------------------------------------------------------------------------------------
'   HasBegin = False
'
'   Call EnableForm(Me, False)
'
'   Call UpdatePartGroupSelected(m_PartGroups)
'
'   Set Ba = New CBalanceAccum
'   Ba.FROM_DATE = uctlFromDate.ShowDate
'   Ba.TO_DATE = uctlToDate.ShowDate
'   Call Ba.ClearData
'   Set Ba = Nothing
'
'   Set BalanceAccums = New Collection
'
'   Set Rs1 = New ADODB.Recordset
'   Set Rs2 = New ADODB.Recordset
'
'   Set TempCol = New Collection
'
'   Set InventoryBals = New Collection
'   Call LoadInventoryBalanceEx(Nothing, BalanceAccums, InternalDateToDate(DateToStringIntLow(uctlFromDate.ShowDate)), uctlToDate.ShowDate, "")
'   Call glbDaily.CopyBalanceAccum(BalanceAccums, InventoryBals)
'
'   Set TempEi = New CLotItem
'
'   Set m_PartItemsDateLocations = Nothing
'   Set m_PartItemsDateLocations = New Collection
''-----------------------------------------------------------------------------------------------------
''                                         Main Operation Here
''-----------------------------------------------------------------------------------------------------
'   NewDate = DateAdd("D", -1, uctlFromDate.ShowDate)
'
'   '=== Detail
'   Set Ei = New CLotItem
'   Ei.LOT_ITEM_ID = -1
'   Ei.FROM_DATE = uctlFromDate.ShowDate
'   Ei.TO_DATE = uctlToDate.ShowDate
'   Ei.COMMIT_FLAG = ""
'   Ei.PIG_FLAG = ""
'   Ei.PART_ITEM_ID = -1
'   Ei.LOCATION_ID = -1
'   Ei.TX_TYPE = "E"
'   Ei.OrderBy = 11
'   Ei.OrderType = 1
''Ei.PART_ITEM_ID = 28
''Ei.LOCATION_ID = 106
'   Call Ei.QueryData(1, Rs1, Count1)
'
'   Set II = New CLotItem
'   II.LOT_ITEM_ID = -1
'   II.FROM_DATE = uctlFromDate.ShowDate
'   II.TO_DATE = uctlToDate.ShowDate
'   II.COMMIT_FLAG = ""
'   II.PIG_FLAG = ""
'   II.PART_ITEM_ID = -1
'   II.LOCATION_ID = -1
'   II.OrderBy = 11
'   II.OrderType = 1
'   II.TX_TYPE = "I"
''II.PART_ITEM_ID = 28
''II.LOCATION_ID = 106
'   Call II.QueryData(1, Rs2, Count2)
'   '== Detail
'
'   Call glbDaily.StartTransaction
'
'   MIN = 0
'   MAX = 100
'   Percent = 0
'   RecordCount = 0
'   prgProgress.MIN = MIN
'   prgProgress.MAX = MAX
'
'   TxCode = "X"
'   While TxCode <> ""
'      Percent = MyDiff(RecordCount, Count1 + Count2) * 100
'      prgProgress.Value = Percent
'      txtPercent.Text = Format(Percent, "0.00")
'
'      TxCode = GetNextTransaction(Rs1, Rs2, II, Ei)
'      If TxCode <> "" Then
'         RecordCount = RecordCount + 1
'
'         If TxCode = "I" Then
'            IsSelectd = IsSelected(II)
'         Else
'            IsSelectd = IsSelected(Ei)
'         End If
'         If Not IsSelectd Then
'            GoTo SkipLabel
'         Else
'            ''Debug.Print
'         End If
'
'         I = I + 1
'         If TxCode = "I" Then
'            '====
''If II.LOT_ITEM_ID = 25688 Then
'''Debug.Print
''End If
'
'            Set O = II
''If DateToStringInt(O.DOCUMENT_DATE) = "2006-02-28 00:00:00" Then
'''Debug.Print
''End If
'            If II.DOCUMENT_TYPE = 3 Then 'ใบโอนวัตถุดิบ
'               Set TempEi = New CLotItem
'               Call GetRelateItem1(O, TempEi)
'               II.INCLUDE_UNIT_PRICE = TempEi.EXPORT_AVG_PRICE
'               II.TOTAL_INCLUDE_PRICE = TempEi.EXPORT_TOTAL_PRICE
'               Set TempEi = Nothing
'            ElseIf II.DOCUMENT_TYPE = 4 Then 'ใบปรับยอดวัตถุดิบ
'            ElseIf II.DOCUMENT_TYPE = 11 Then 'ใบสั่งผลิต
'               Call CalculateRMPrice(II.INVENTORY_DOC_ID, II, InventoryBals)
'            Else
'               II.INCLUDE_UNIT_PRICE = MyDiffEx(II.TOTAL_INCLUDE_PRICE, II.IMPORT_AMOUNT)
'            End If
'            '====
'         ElseIf TxCode = "E" Then
'            Set O = Ei
'         End If
'
'         TempKey = O.LOCATION_ID & "-" & O.PART_ITEM_ID
'
'         Set TempII = GetImportItem(InventoryBals, TempKey)
'         If TempII.PART_ITEM_ID <= 0 Then
'            'Get balance item here
'            Set TempO = GetImportItem(InventoryBals, TempKey)
'
'            Set TempII = New CLotItem
'            TempII.LOCATION_ID = O.LOCATION_ID
'            TempII.PART_ITEM_ID = O.PART_ITEM_ID
'            If O.TX_TYPE = "I" Then
'               TempII.INCLUDE_UNIT_PRICE = MyDiffEx(O.TOTAL_INCLUDE_PRICE, O.IMPORT_AMOUNT)
'               TempII.CURRENT_AMOUNT = O.IMPORT_AMOUNT
'               TempII.TOTAL_INCLUDE_PRICE = O.TOTAL_INCLUDE_PRICE
'            ElseIf O.TX_TYPE = "E" Then
'               TempII.INCLUDE_UNIT_PRICE = O.EXPORT_AVG_PRICE
'               TempII.CURRENT_AMOUNT = -1 * O.EXPORT_AMOUNT
'               TempII.TOTAL_INCLUDE_PRICE = O.EXPORT_TOTAL_PRICE
'            End If
'
'            Call InventoryBals.add(TempII, TempKey)
'            Set TempII = Nothing
'            Set TempII = GetImportItem(InventoryBals, TempKey)
'         Else
''If O.DOCUMENT_NO = "ADJ 4-2" And (O.PART_ITEM_ID = 28) And (O.LOCATION_ID = 106) Then
'''Debug.Print
''End If
'            If O.TX_TYPE = "I" Then
'               If O.ADJUST_FLAG = "Y" Then
'                  TempII.CURRENT_AMOUNT = O.NEED_TOTAL_AMOUNT   'Val(Format(TempII.CURRENT_AMOUNT + O.IMPORT_AMOUNT, "0.00000"))
'                  TempII.TOTAL_INCLUDE_PRICE = O.NEED_TOTAL_PRICE   'TempII.TOTAL_INCLUDE_PRICE + O.TOTAL_INCLUDE_PRICE
'                  If TempII.CURRENT_AMOUNT > 0 Then
'                     TempII.INCLUDE_UNIT_PRICE = MyDiffEx(TempII.TOTAL_INCLUDE_PRICE + O.TOTAL_INCLUDE_PRICE, TempII.CURRENT_AMOUNT + O.IMPORT_AMOUNT)   'TempO.NEW_PRICE
'                  Else
'                     TempII.INCLUDE_UNIT_PRICE = O.NEED_AVG_PRICE
'                     'ใช้ค่าเดิม
'                  End If
'               Else
'                  TempII.INCLUDE_UNIT_PRICE = MyDiffEx(TempII.TOTAL_INCLUDE_PRICE + O.TOTAL_INCLUDE_PRICE, Val(Format(TempII.CURRENT_AMOUNT + O.IMPORT_AMOUNT, "0.00000")))   'TempO.NEW_PRICE
'                  TempII.CURRENT_AMOUNT = Val(Format(TempII.CURRENT_AMOUNT + O.IMPORT_AMOUNT, "0.00000"))
'                  TempII.TOTAL_INCLUDE_PRICE = TempII.TOTAL_INCLUDE_PRICE + O.TOTAL_INCLUDE_PRICE
'               End If
'            ElseIf O.TX_TYPE = "E" Then
'               TempII.CURRENT_AMOUNT = TempII.CURRENT_AMOUNT - O.EXPORT_AMOUNT 'Val(Format(TempII.CURRENT_AMOUNT - O.EXPORT_AMOUNT, "0.00000"))
'               If O.ADJUST_FLAG = "Y" Then
'                  TempII.TOTAL_INCLUDE_PRICE = TempII.NEED_TOTAL_PRICE   'TempII.TOTAL_INCLUDE_PRICE - O.TOTAL_INCLUDE_PRICE
'                  TempII.CURRENT_AMOUNT = TempII.NEED_TOTAL_AMOUNT
'               Else
'                  TempII.TOTAL_INCLUDE_PRICE = TempII.TOTAL_INCLUDE_PRICE - (TempII.INCLUDE_UNIT_PRICE * O.EXPORT_AMOUNT)
'               End If
'            End If
'         End If
'
'         Call GeneratePartItemLocationDate(O, TempII)
'
'         If TxCode = "I" Then
'            PrevAmount = TempII.CURRENT_AMOUNT - O.IMPORT_AMOUNT
'            CurrentAmount = PrevAmount + II.IMPORT_AMOUNT
''            CurrentAmount = Format(CurrentAmount, "0.00000000")
'            If CurrentAmount > 0 Then
'               AvgPrice = Val(Format(MyDiffEx(TempII.TOTAL_INCLUDE_PRICE, CurrentAmount), "0.00000"))
'            Else
'               AvgPrice = Val(Format(TempII.INCLUDE_UNIT_PRICE, "0.00000"))
'            End If
'            NewTotalPrice = TempII.TOTAL_INCLUDE_PRICE
''''Debug.Print "I " & " " & II.DOCUMENT_NO & " " & PrevAmount & " " & II.IMPORT_AMOUNT & " " & CurrentAmount & " " & AvgPrice & " " & NewTotalPrice
'            Call II.PatchAvgPrice(II.INCLUDE_UNIT_PRICE, PrevAmount, CurrentAmount, AvgPrice, II.IMPORT_AMOUNT, II.DOCUMENT_TYPE, II.TOTAL_INCLUDE_PRICE, NewTotalPrice)
'         ElseIf TxCode = "E" Then
'            PrevAmount = TempII.CURRENT_AMOUNT + Ei.EXPORT_AMOUNT
'            CurrentAmount = PrevAmount - Ei.EXPORT_AMOUNT
'            If O.ADJUST_FLAG = "Y" Then
'               AvgPrice = TempII.INCLUDE_UNIT_PRICE 'MyDiffEx(TempII.TOTAL_INCLUDE_PRICE, CurrentAmount) 'เดิมทีไม่มี -1 คูณ
'               NewTotalPrice = TempII.TOTAL_INCLUDE_PRICE  'CurrentAmount * AvgPrice
'               ExportTotalPrice = O.TOTAL_INCLUDE_PRICE
'            Else
'               AvgPrice = TempII.INCLUDE_UNIT_PRICE
'               NewTotalPrice = CurrentAmount * AvgPrice
'               ExportTotalPrice = AvgPrice * Ei.EXPORT_AMOUNT     'TOTAL_INCLUDE_PRICE
'            End If
'            NewTotalPrice = Val(Format(NewTotalPrice, "0.00000"))
''''Debug.Print "E " & " " & Ei.DOCUMENT_NO & " " & PrevAmount & " " & Ei.EXPORT_AMOUNT & " " & CurrentAmount & " " & AvgPrice & " " & NewTotalPrice
'            Call Ei.PatchAvgPriceExp(AvgPrice, PrevAmount, CurrentAmount, ExportTotalPrice, NewTotalPrice)
'         End If
'      End If 'Tx code
'      DoEvents
'
'SkipLabel:
'   Wend
'
'   'Call InsertBalanceAccum
'
'   txtPercent.Text = Format(100, "0.00")
'   prgProgress.Value = 100
'   Call glbDaily.CommitTransaction
'   HasBegin = False
'
'   If Rs1.State = adStateOpen Then
'      Rs1.Close
'   End If
'   Set Rs1 = Nothing
'
'   If Rs2.State = adStateOpen Then
'      Rs2.Close
'   End If
'   Set Rs2 = Nothing
'
'   Set Ei = Nothing
'   Set II = Nothing
'   Set TempEi = Nothing
'   Set InventoryBals = Nothing
'   Set BalanceAccums = Nothing
'   Set TempCol = Nothing
'   Call EnableForm(Me, True)
'
'   Exit Sub
'
''ErrHandler:
''   If HasBegin Then
''      glbDaily.RollbackTransaction
''   End If
''   glbErrorLog.LocalErrorMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
''   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'End Sub

Private Sub ReformatAdjust(InventoryBals As Collection, O As CLotItem, NewTxCode As String)
Dim TempII As CLotItem
Dim TempKey As String
Dim TxAmount As Double
Dim Mult As Long

   TempKey = O.LOCATION_ID & "-" & O.PART_ITEM_ID
   Set TempII = GetImportItem(InventoryBals, TempKey)

   TxAmount = O.NEED_TOTAL_AMOUNT - TempII.CURRENT_AMOUNT
   If TxAmount >= 0 Then
      NewTxCode = "I"
      O.IMPORT_AMOUNT = Abs(TxAmount)
      Mult = 1
   Else
      NewTxCode = "E"
      O.EXPORT_AMOUNT = Abs(TxAmount)
      Mult = -1
   End If
   
   O.TX_TYPE = NewTxCode
   O.TX_AMOUNT = Abs(TxAmount)
   If O.AUTO_PRICE = "Y" Then
      'ให้ระบบคำนวณราคาให้
      O.NEED_TOTAL_PRICE = O.NEED_TOTAL_AMOUNT * MyDiffEx(TempII.TOTAL_INCLUDE_PRICE, TempII.CURRENT_AMOUNT)
      O.NEW_PRICE = MyDiffEx(TempII.TOTAL_INCLUDE_PRICE, TempII.CURRENT_AMOUNT)
   End If
   O.TOTAL_INCLUDE_PRICE = Mult * (O.NEED_TOTAL_PRICE - TempII.TOTAL_INCLUDE_PRICE)
   O.TOTAL_ACTUAL_PRICE = O.TOTAL_INCLUDE_PRICE
   O.INCLUDE_UNIT_PRICE = MyDiffEx(O.TOTAL_INCLUDE_PRICE, O.TX_AMOUNT)
   O.ACTUAL_UNIT_PRICE = O.INCLUDE_UNIT_PRICE
End Sub

Private Sub cmdStart_Click()
On Error GoTo ErrHandler
Dim Percent As Double
Dim MIN As Double
Dim MAX As Double
Dim RecordCount As Double
Dim O As Object
Dim TempO As Object
Dim InventoryBals As Collection
Dim RName As String
Dim cData As CPartLocation
Dim I As Long
Dim J As Long
Dim strFormat As String
Dim IsOK As Boolean
Dim Amt As Double
Dim Ei As CLotItem
Dim II As CLotItem
Dim Ai As CLotItem
Dim Rs1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Rs3 As ADODB.Recordset
Dim TxCode As String
Dim iCount As Long
Dim AvgPrice As Double
Dim PrevAmount As Double
Dim PrevPrice As Double
Dim CurrentAmount As Double
Dim HasBegin As Boolean
Dim TempII As CLotItem
Dim TempKey As String
Dim Count1 As Long
Dim Count2 As Long
Dim Count3 As Long
Dim TempCol As Collection
Dim TempEi As CLotItem
Dim ExportTotalPrice As Double
Dim NewDate As Date
Dim Ba As CBalanceAccum
Dim BalanceAccums As Collection
Dim NewTotalPrice As Double
Dim IsSelectd As Boolean
Dim NewTxCode As String
Dim MaColls As Collection
Dim MonthlyAccums  As Collection
Dim FromDate As Date
Dim ToDate As Date
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   HasBegin = False
   
   Call EnableForm(Me, False)
   
   Call UpdatePartGroupSelected(m_PartGroups)
   
   Call GetFirstLastDate(DateSerial(Val(txtYear.Text) - 543, cboMonth.ItemData(Minus2Zero(cboMonth.ListIndex)), 1), FromDate, ToDate)
   Set Ba = New CBalanceAccum
   Ba.FROM_DATE = FromDate
   Ba.TO_DATE = ToDate
   Call Ba.ClearData
   Set Ba = Nothing
   
   Set BalanceAccums = New Collection
   
   Set Rs1 = New ADODB.Recordset
   Set Rs2 = New ADODB.Recordset
   Set Rs3 = New ADODB.Recordset
   
   Set TempCol = New Collection
   
   Set InventoryBals = New Collection
   
   If chkCheckBalance.Value = ssCBChecked Then
      Call LoadInventoryBalanceEx(Nothing, BalanceAccums, InternalDateToDate(DateToStringIntLow(FromDate)), ToDate, "")
      Call glbDaily.CopyBalanceAccum(BalanceAccums, InventoryBals)
      Call glbDaily.StartTransaction
      HasBegin = True
      Set MaColls = New Collection
      Call InsertMonthlyAccum(InventoryBals, MaColls, DateAdd("D", -1, FromDate))       'แปลง จากยอดยกมาเป็น Monthly Collection
      Call InsertBalanceAccum(MaColls)             'เพิ่มข้อมูลของ Monthly Collection จากข้อมูลเคลื่อนไหวเดือนปัจจุบัน
      
      Set MaColls = Nothing
      txtPercent.Text = Format(100, "0.00")
      prgProgress.Value = 100
      Call glbDaily.CommitTransaction
      HasBegin = False
      Call EnableForm(Me, True)
      Exit Sub
   Else
      Dim YYYYMM   As String
      YYYYMM = Format(Year(DateAdd("D", -1, FromDate)), "0000") & "-" & Format(Month(DateAdd("D", -1, FromDate)), "00")
      Set MonthlyAccums = New Collection
      Call LoadMonthlyBalance(Nothing, MonthlyAccums, YYYYMM)
      Call glbDaily.CopyMonthlyAccum(MonthlyAccums, InventoryBals)
   End If
   
   Dim Ma As CMonthlyAccum
   Set Ma = New CMonthlyAccum
   Ma.FROM_YYYYMM = Format(Year(FromDate), "0000") & "-" & Format(Month(FromDate), "00")
   Ma.TO_YYYYMM = Format(Year(ToDate), "0000") & "-" & Format(Month(ToDate), "00")
   Call Ma.ClearData
   
   Set TempEi = New CLotItem
   
   Set m_PartItemsDateLocations = Nothing
   Set m_PartItemsDateLocations = New Collection
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   NewDate = DateAdd("D", -1, FromDate)
   
   '=== Detail
   Set Ei = New CLotItem
   Ei.LOT_ITEM_ID = -1
   Ei.FROM_DATE = FromDate
   Ei.TO_DATE = ToDate
   Ei.COMMIT_FLAG = ""
   Ei.PIG_FLAG = ""
'                                                                                                                     Ei.PART_ITEM_ID = 58
'                                                                                                                     Ei.LOCATION_ID = 109
'                                                                                                                     Ei.PART_ITEM_ID = 3199
'                                                                                                                     Ei.LOCATION_ID = 110
   Ei.TX_TYPE = "E"
   Ei.OrderBy = 11
   Ei.OrderType = 1
'   Ei.PART_ITEM_ID = -1
'   Ei.LOCATION_ID = -1
   Call Ei.QueryData(1, Rs1, Count1, True)
   
   Set II = New CLotItem
   II.LOT_ITEM_ID = -1
   II.FROM_DATE = FromDate
   II.TO_DATE = ToDate
   II.COMMIT_FLAG = ""
   II.PIG_FLAG = ""
'                                                                                                                        II.PART_ITEM_ID = 58
'                                                                                                                        II.LOCATION_ID = 109
'                                                                                                                     II.PART_ITEM_ID = 3199
'                                                                                                                     II.LOCATION_ID = 110
   II.OrderBy = 11
   II.OrderType = 1
   II.TX_TYPE = "I"
'   II.PART_ITEM_ID = -1
'   II.LOCATION_ID = -1
   Call II.QueryData(1, Rs2, Count2, True)
   '== Detail

   Call glbDaily.StartTransaction
   HasBegin = True
   MIN = 0
   MAX = 100
   Percent = 0
   RecordCount = 0
   prgProgress.MIN = MIN
   prgProgress.MAX = MAX
   
   TxCode = "X"
   While TxCode <> ""
     DoEvents
      Percent = MyDiff(RecordCount, Count1 + Count2) * 100
      prgProgress.Value = Percent
      txtPercent.Text = Format(Percent, "0.00")
      
      TxCode = GetNextTransaction(Rs1, Rs2, II, Ei)
      If TxCode <> "" Then
         RecordCount = RecordCount + 1
         
         If TxCode = "I" Then
            IsSelectd = IsSelected(II)
         Else
            IsSelectd = IsSelected(Ei)
         End If
         If Not IsSelectd Then
            GoTo SkipLabel
         End If

         If (TxCode = "I") Then
            Set O = II
         ElseIf TxCode = "E" Then
            Set O = Ei
         End If
'If (O.DOCUMENT_NO = "FD-สำเร็จรูป-010761" And O.PART_ITEM_ID = 4836) Then
'   'Debug.Print
'End If

'If (O.DOCUMENT_NO = "JP-050034865") Then
'   'Debug.Print
'End If



'      ''Debug.Print (InventoryBals.Count)
         If (O.DOCUMENT_TYPE = 4) Or (O.DOCUMENT_TYPE = 5) Then 'ใบปรับยอด ถ้าเป็น adjust ตอนแรกจะเป็น I
'If (O.DOCUMENT_TYPE = 4) And (O.PART_ITEM_ID = 1210) Then
''Debug.Print
'End If
            Call ReformatAdjust(InventoryBals, O, NewTxCode)
            If NewTxCode = "E" Then
               'เสมือนได้ เรคคอร์ดที่เป็น export ออกมา
               Ei.TX_TYPE = "E"
               Ei.INCLUDE_UNIT_PRICE = O.INCLUDE_UNIT_PRICE
               Ei.TOTAL_INCLUDE_PRICE = O.TOTAL_INCLUDE_PRICE
               Ei.ACTUAL_UNIT_PRICE = O.ACTUAL_UNIT_PRICE
               Ei.TOTAL_ACTUAL_PRICE = O.TOTAL_ACTUAL_PRICE
               Ei.TX_AMOUNT = O.TX_AMOUNT
               Ei.EXPORT_AMOUNT = O.EXPORT_AMOUNT
               Ei.LOCATION_ID = O.LOCATION_ID
               Ei.PART_ITEM_ID = O.PART_ITEM_ID
               Ei.ADJUST_FLAG = O.ADJUST_FLAG
               Ei.NEED_TOTAL_AMOUNT = O.NEED_TOTAL_AMOUNT
               Ei.NEED_TOTAL_PRICE = O.NEED_TOTAL_PRICE
               Ei.NEED_AVG_PRICE = O.NEED_AVG_PRICE
               Ei.LOT_ITEM_ID = O.LOT_ITEM_ID
               Ei.DOCUMENT_NO = O.DOCUMENT_NO
               Ei.NEW_PRICE = O.NEW_PRICE
If O.AUTO_PRICE = "N" Then
Ei.NEW_PRICE = MyDiffEx(Ei.NEED_TOTAL_PRICE, Ei.NEED_TOTAL_AMOUNT)
End If
            ElseIf NewTxCode = "I" Then
               II.TX_TYPE = "I"
               II.INCLUDE_UNIT_PRICE = O.INCLUDE_UNIT_PRICE
               II.TOTAL_INCLUDE_PRICE = O.TOTAL_INCLUDE_PRICE
               II.ACTUAL_UNIT_PRICE = O.ACTUAL_UNIT_PRICE
               II.TOTAL_ACTUAL_PRICE = O.TOTAL_ACTUAL_PRICE
               II.TX_AMOUNT = O.TX_AMOUNT
               II.IMPORT_AMOUNT = II.TX_AMOUNT
               II.LOCATION_ID = O.LOCATION_ID
               II.PART_ITEM_ID = O.PART_ITEM_ID
               II.ADJUST_FLAG = O.ADJUST_FLAG
               II.NEED_TOTAL_AMOUNT = O.NEED_TOTAL_AMOUNT
               II.NEED_TOTAL_PRICE = O.NEED_TOTAL_PRICE
               II.NEED_AVG_PRICE = O.NEED_AVG_PRICE
               II.LOT_ITEM_ID = O.LOT_ITEM_ID
               II.DOCUMENT_NO = O.DOCUMENT_NO
               II.DOCUMENT_TYPE = O.DOCUMENT_TYPE
               II.MANUAL_PRICE = O.MANUAL_PRICE
            End If
            TxCode = NewTxCode
         End If
         
         I = I + 1
         If TxCode = "I" Then
            '====
            Set O = II
            
            If (II.DOCUMENT_TYPE = 3) Or (II.DOCUMENT_TYPE = 22) Then 'ใบโอนวัตถุดิบ
               Set TempEi = New CLotItem
               Call GetRelateItem1(O, TempEi)
               II.INCLUDE_UNIT_PRICE = TempEi.EXPORT_AVG_PRICE
               II.TOTAL_INCLUDE_PRICE = TempEi.EXPORT_TOTAL_PRICE
               Set TempEi = Nothing
            ElseIf (II.DOCUMENT_TYPE = 12) Or (II.DOCUMENT_TYPE = 13) Or (II.DOCUMENT_TYPE = 14) Then 'ใบสั่งผลิต เปลี่ยนเป็น 12 13 14 แทน ของเดิมเป็น 11
'If (O.DOCUMENT_NO = "P-00758") Then
''Debug.Print
'End If
               Call CalculateRMPrice(II.INVENTORY_DOC_ID, II, InventoryBals)
            Else
               II.INCLUDE_UNIT_PRICE = MyDiffEx(II.TOTAL_INCLUDE_PRICE, II.IMPORT_AMOUNT)
            End If
            '====
         ElseIf TxCode = "E" Then
'If (O.DOCUMENT_NO = "025-1202-9") Then
''Debug.Print
'End If
            Set O = Ei
'If (Ei.DOCUMENT_TYPE = 4) And (Ei.PART_ITEM_ID = 3) And (Ei.DOCUMENT_NO = "BIN010549-1") Then
''   'Debug.Print
'End If
         End If

         TempKey = O.LOCATION_ID & "-" & O.PART_ITEM_ID

         Set TempII = GetImportItem(InventoryBals, TempKey)
         If TempII.PART_ITEM_ID <= 0 Then
            'Get balance item here
            Set TempO = GetImportItem(InventoryBals, TempKey)

            Set TempII = New CLotItem
            TempII.LOCATION_ID = O.LOCATION_ID
            TempII.PART_ITEM_ID = O.PART_ITEM_ID
            If O.TX_TYPE = "I" Then
               TempII.INCLUDE_UNIT_PRICE = MyDiffEx(O.TOTAL_INCLUDE_PRICE, O.IMPORT_AMOUNT)
               TempII.CURRENT_AMOUNT = O.IMPORT_AMOUNT
               TempII.TOTAL_INCLUDE_PRICE = O.TOTAL_INCLUDE_PRICE
            ElseIf O.TX_TYPE = "E" Then
               TempII.INCLUDE_UNIT_PRICE = O.EXPORT_AVG_PRICE
               TempII.CURRENT_AMOUNT = -1 * O.EXPORT_AMOUNT
               TempII.TOTAL_INCLUDE_PRICE = O.EXPORT_TOTAL_PRICE
            End If

            Call InventoryBals.add(TempII, TempKey)
            Set TempII = Nothing
            Set TempII = GetImportItem(InventoryBals, TempKey)
         Else
            If O.TX_TYPE = "I" Then
               If O.ADJUST_FLAG = "Y" Then
                  TempII.CURRENT_AMOUNT = O.NEED_TOTAL_AMOUNT   'Val(Format(TempII.CURRENT_AMOUNT + O.IMPORT_AMOUNT, "0.00000"))
                  TempII.TOTAL_INCLUDE_PRICE = O.NEED_TOTAL_PRICE   'TempII.TOTAL_INCLUDE_PRICE + O.TOTAL_INCLUDE_PRICE
                  If TempII.CURRENT_AMOUNT > 0 Then
                     TempII.INCLUDE_UNIT_PRICE = MyDiffEx(TempII.TOTAL_INCLUDE_PRICE, TempII.CURRENT_AMOUNT)
                     'MyDiffEx(TempII.TOTAL_INCLUDE_PRICE + O.TOTAL_INCLUDE_PRICE, TempII.CURRENT_AMOUNT + O.IMPORT_AMOUNT)   'TempO.NEW_PRICE
                  Else
                     TempII.INCLUDE_UNIT_PRICE = O.NEED_AVG_PRICE
                     'ใช้ค่าเดิม
                  End If
               Else
                  TempII.INCLUDE_UNIT_PRICE = MyDiffEx(TempII.TOTAL_INCLUDE_PRICE + O.TOTAL_INCLUDE_PRICE, Val(Format(TempII.CURRENT_AMOUNT + O.IMPORT_AMOUNT, "0.00000")))   'TempO.NEW_PRICE
                  TempII.CURRENT_AMOUNT = Val(Format(TempII.CURRENT_AMOUNT + O.IMPORT_AMOUNT, "0.00000"))
                  TempII.TOTAL_INCLUDE_PRICE = TempII.TOTAL_INCLUDE_PRICE + O.TOTAL_INCLUDE_PRICE
               End If
            ElseIf O.TX_TYPE = "E" Then
               TempII.CURRENT_AMOUNT = TempII.CURRENT_AMOUNT - O.EXPORT_AMOUNT
               If O.ADJUST_FLAG = "Y" Then
                  TempII.TOTAL_INCLUDE_PRICE = O.NEED_TOTAL_PRICE
                  TempII.CURRENT_AMOUNT = O.NEED_TOTAL_AMOUNT
                  If O.AUTO_PRICE = "N" Then
                     TempII.INCLUDE_UNIT_PRICE = O.NEED_AVG_PRICE
                  Else
                     'ไม่ต้องทำอะไร ใช้ราคาของระบบ
                  End If
               Else
                  TempII.TOTAL_INCLUDE_PRICE = TempII.TOTAL_INCLUDE_PRICE - (TempII.INCLUDE_UNIT_PRICE * O.EXPORT_AMOUNT)
               End If
            End If
         End If

         Call GeneratePartItemLocationDate(O, TempII)

         If TxCode = "I" Then
            PrevAmount = TempII.CURRENT_AMOUNT - O.IMPORT_AMOUNT
            CurrentAmount = PrevAmount + II.IMPORT_AMOUNT
            If CurrentAmount > 0 Then
               AvgPrice = Val(Format(MyDiffEx(TempII.TOTAL_INCLUDE_PRICE, CurrentAmount), "0.00000"))
            Else
               AvgPrice = Val(Format(TempII.INCLUDE_UNIT_PRICE, "0.00000"))
            End If
            NewTotalPrice = TempII.TOTAL_INCLUDE_PRICE
            Call II.PatchAvgPrice(II.INCLUDE_UNIT_PRICE, PrevAmount, CurrentAmount, AvgPrice, II.IMPORT_AMOUNT, II.DOCUMENT_TYPE, II.TOTAL_INCLUDE_PRICE, NewTotalPrice)
         ElseIf TxCode = "E" Then
            PrevAmount = TempII.CURRENT_AMOUNT + Ei.EXPORT_AMOUNT
            CurrentAmount = PrevAmount - Ei.EXPORT_AMOUNT
            If O.ADJUST_FLAG = "Y" Then
               AvgPrice = TempII.INCLUDE_UNIT_PRICE
'               If O.AUTO_PRICE = "Y" Then
'                  PrevPrice = AvgPrice
'               Else
'                  PrevPrice = Ei.PREVIOUS_PRICE
'               End If
               NewTotalPrice = TempII.TOTAL_INCLUDE_PRICE
               ExportTotalPrice = O.TOTAL_INCLUDE_PRICE
            Else
               AvgPrice = TempII.INCLUDE_UNIT_PRICE
               PrevPrice = TempII.INCLUDE_UNIT_PRICE
               NewTotalPrice = TempII.TOTAL_INCLUDE_PRICE  'CurrentAmount * AvgPrice
               ExportTotalPrice = AvgPrice * Ei.EXPORT_AMOUNT
            End If
            NewTotalPrice = Val(Format(NewTotalPrice, "0.00000"))
            Call Ei.PatchAvgPriceExp(AvgPrice, PrevAmount, CurrentAmount, ExportTotalPrice, NewTotalPrice)
         End If
      End If 'Tx code
      DoEvents
      
SkipLabel:
   Wend
   
   Set MaColls = New Collection
   Call InsertMonthlyAccum(InventoryBals, MaColls, FromDate)
   Call InsertBalanceAccum(MaColls)
   
   Set MaColls = Nothing
   txtPercent.Text = Format(100, "0.00")
   prgProgress.Value = 100
   Call glbDaily.CommitTransaction
   HasBegin = False
   
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
      
   If Rs2.State = adStateOpen Then
      Rs2.Close
   End If
   Set Rs2 = Nothing
      
   Set Ei = Nothing
   Set II = Nothing
   Set TempEi = Nothing
   Set InventoryBals = Nothing
   Set BalanceAccums = Nothing
   Set TempCol = Nothing
   Call EnableForm(Me, True)
   
   Exit Sub
   
ErrHandler:
   If HasBegin Then
      glbDaily.RollbackTransaction
   End If
   glbErrorLog.LocalErrorMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub InsertBalanceAccum(MaColls As Collection)
'On Error Resume Next
Dim Ba As CBalanceAccum
Dim II As CLotItem
Dim iCount As Long
Dim TempMa As CMonthlyAccum
   
   For Each II In m_PartItemsDateLocations
      Set Ba = New CBalanceAccum
      '''Debug.Print DateToStringInt(II.DOCUMENT_DATE) & " "; II.INCLUDE_UNIT_PRICE & " " & II.ALL_IMPORT_AMT & " " & II.ALL_EXPORT_AMT & " " & II.BALANCE_AMOUNT
      
      Set TempMa = GetObject("CMonthlyAccum", MaColls, Trim(II.PART_ITEM_ID & "-" & II.LOCATION_ID & "-" & Format(Year(II.DOCUMENT_DATE), "0000") & "-" & Format(Month(II.DOCUMENT_DATE), "00")), False)
      If TempMa Is Nothing Then
         Set TempMa = New CMonthlyAccum
         TempMa.PART_ITEM_ID = II.PART_ITEM_ID
         TempMa.LOCATION_ID = II.LOCATION_ID
         TempMa.YYYYMM = Format(Year(II.DOCUMENT_DATE), "0000") & "-" & Format(Month(II.DOCUMENT_DATE), "00")
         Call MaColls.add(TempMa, Trim(II.PART_ITEM_ID & "-" & II.LOCATION_ID & "-" & Format(Year(II.DOCUMENT_DATE), "0000") & "-" & Format(Month(II.DOCUMENT_DATE), "00")))
      End If
      TempMa.BALANCE_AMOUNT = II.BALANCE_AMOUNT
      TempMa.TOTAL_INCLUDE_PRICE = II.TOTAL_INCLUDE_PRICE
      TempMa.AVG_PRICE = II.INCLUDE_UNIT_PRICE
      
      
      Ba.PART_ITEM_ID = II.PART_ITEM_ID
      Ba.FROM_DATE = II.DOCUMENT_DATE
      Ba.TO_DATE = II.DOCUMENT_DATE
      Ba.LOCATION_ID = II.LOCATION_ID
      'Call Ba.QueryData(1, m_Rs, iCount)
      'If m_Rs.EOF Then
         Ba.AddEditMode = SHOW_ADD
      'Else
       '  Call Ba.PopulateFromRS(1, m_Rs)
        ' Ba.AddEditMode = SHOW_EDIT
      'End If
      Ba.DOCUMENT_DATE = II.DOCUMENT_DATE
      Ba.IMPORT_AMOUNT = II.ALL_IMPORT_AMT
      Ba.EXPORT_AMOUNT = II.ALL_EXPORT_AMT
      Ba.BALANCE_AMOUNT = II.BALANCE_AMOUNT
      Ba.TOTAL_INCLUDE_PRICE = II.TOTAL_INCLUDE_PRICE
      Ba.AVG_PRICE = II.INCLUDE_UNIT_PRICE
      Call Ba.AddEditData
      
      Set Ba = Nothing
   Next II
   
   For Each TempMa In MaColls
      TempMa.AddEditMode = SHOW_ADD
      Call TempMa.AddEditData
   Next TempMa
End Sub
Private Sub InsertMonthlyAccum(Src As Collection, Des As Collection, DateInsert As Date)
'On Error Resume Next
Dim Ma As CMonthlyAccum
Dim II As CLotItem
Dim iCount As Long
   
   For Each II In Src
      If II.CURRENT_AMOUNT <> 0 Then
         Set Ma = New CMonthlyAccum
         
         If II.PART_ITEM_ID = 4836 Then
            'Debug.Print
         End If
         
         Ma.PART_ITEM_ID = II.PART_ITEM_ID
         Ma.YYYYMM = Format(Year(DateInsert), "0000") & "-" & Format(Month(DateInsert), "00")
         Ma.LOCATION_ID = II.LOCATION_ID
         
         Ma.BALANCE_AMOUNT = II.CURRENT_AMOUNT
         Ma.TOTAL_INCLUDE_PRICE = II.TOTAL_INCLUDE_PRICE
         Ma.AVG_PRICE = II.INCLUDE_UNIT_PRICE
               
         Call Des.add(Ma, Trim(II.PART_ITEM_ID & "-" & II.LOCATION_ID & "-" & Ma.YYYYMM))
         Set Ma = Nothing
      End If
   Next II
   
End Sub

Private Sub Form_Activate()

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      cboMonth.ListIndex = IDToListIndex(cboMonth, Month(Now))
      txtYear.Text = Val(Year(Now)) + 543
      
      Call LoadPartGroup(Nothing, m_PartGroups)
      Call LoadPartGroupView(m_PartGroups)

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

Private Sub LoadPartGroupView(Col As Collection)
Dim C As CPartGroup
Dim N As Node
Dim Np As Node

      For Each C In Col
         Set N = TreeView1.Nodes.add(, tvwFirst, Trim(str(C.PART_GROUP_ID)) & "-X", C.PART_GROUP_NAME & " (" & C.PART_GROUP_NO & ")", 1, 1)
         N.Tag = C.PART_GROUP_ID
         N.Checked = False
         
         N.Expanded = False
      Next C
End Sub

Private Sub UpdatePartGroupSelected(Col As Collection)
Dim C As CPartGroup
Dim N As Node
Dim Count As Long

   Count = 0
   For Each N In TreeView1.Nodes
      Set C = GetPartGroup(m_PartGroups, N.Tag)
      If N.Checked Then
         C.SELECT_FLAG = "Y"
         Count = Count + 1
      Else
         C.SELECT_FLAG = "N"
      End If
   Next N
   
   If Count <= 0 Then
      For Each C In m_PartGroups
         C.SELECT_FLAG = "Y"
      Next C
   End If
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = "ปรับราคาเฉลี่ย"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblFileName, "เดือนปี")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblPartGroup, "กลุ่มวัตถุดิบ")

'   Call InitCheckBox(chkBalanceFlag, "ลบยอดยกมา")
'   chkBalanceFlag.Value = ssCBUnchecked
   
   Call InitCheckBox(chkCheckBalance, "คำนวณยอดยกมาใหม่")
   Call InitCombo(cboMonth)
   Call InitThaiMonth(cboMonth)

   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False

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
   Set m_Balances = New Collection
   Set m_PartItemsDateLocations = New Collection
   Set m_PartGroups = New Collection
   
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
   Set m_Balances = Nothing
   Set m_PartItemsDateLocations = Nothing
   Set m_PartGroups = Nothing
End Sub
