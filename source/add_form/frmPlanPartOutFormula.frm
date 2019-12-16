VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPlanPartOutFormula 
   ClientHeight    =   10770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14760
   Icon            =   "frmPlanPartOutFormula.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10770
   ScaleWidth      =   14760
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   10800
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   19050
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlDate uctlToDate 
         Height          =   405
         Left            =   5340
         TabIndex        =   1
         Top             =   870
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2865
         Left            =   270
         TabIndex        =   3
         Top             =   1530
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   5054
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
         Column(1)       =   "frmPlanPartOutFormula.frx":27A2
         Column(2)       =   "frmPlanPartOutFormula.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmPlanPartOutFormula.frx":290E
         FormatStyle(2)  =   "frmPlanPartOutFormula.frx":2A6A
         FormatStyle(3)  =   "frmPlanPartOutFormula.frx":2B1A
         FormatStyle(4)  =   "frmPlanPartOutFormula.frx":2BCE
         FormatStyle(5)  =   "frmPlanPartOutFormula.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmPlanPartOutFormula.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   14805
         _ExtentX        =   26114
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1380
         TabIndex        =   0
         Top             =   870
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin GridEX20.GridEX GridEX2 
         Height          =   2865
         Left            =   7920
         TabIndex        =   5
         Top             =   1530
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   5054
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
         Column(1)       =   "frmPlanPartOutFormula.frx":2F36
         Column(2)       =   "frmPlanPartOutFormula.frx":2FFE
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmPlanPartOutFormula.frx":30A2
         FormatStyle(2)  =   "frmPlanPartOutFormula.frx":31FE
         FormatStyle(3)  =   "frmPlanPartOutFormula.frx":32AE
         FormatStyle(4)  =   "frmPlanPartOutFormula.frx":3362
         FormatStyle(5)  =   "frmPlanPartOutFormula.frx":343A
         ImageCount      =   0
         PrinterProperties=   "frmPlanPartOutFormula.frx":34F2
      End
      Begin GridEX20.GridEX GridEX3 
         Height          =   5265
         Left            =   7920
         TabIndex        =   11
         Top             =   4560
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   9287
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
         Column(1)       =   "frmPlanPartOutFormula.frx":36CA
         Column(2)       =   "frmPlanPartOutFormula.frx":3792
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmPlanPartOutFormula.frx":3836
         FormatStyle(2)  =   "frmPlanPartOutFormula.frx":3992
         FormatStyle(3)  =   "frmPlanPartOutFormula.frx":3A42
         FormatStyle(4)  =   "frmPlanPartOutFormula.frx":3AF6
         FormatStyle(5)  =   "frmPlanPartOutFormula.frx":3BCE
         ImageCount      =   0
         PrinterProperties=   "frmPlanPartOutFormula.frx":3C86
      End
      Begin prjFarmManagement.uctlDate uctlPlanDate 
         Height          =   405
         Left            =   2640
         TabIndex        =   12
         Top             =   10080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtFormulaNo 
         Height          =   435
         Left            =   10620
         TabIndex        =   14
         Top             =   870
         Width           =   1935
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX4 
         Height          =   5265
         Left            =   240
         TabIndex        =   16
         Top             =   4560
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   9287
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
         Column(1)       =   "frmPlanPartOutFormula.frx":3E5E
         Column(2)       =   "frmPlanPartOutFormula.frx":3F26
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmPlanPartOutFormula.frx":3FCA
         FormatStyle(2)  =   "frmPlanPartOutFormula.frx":4126
         FormatStyle(3)  =   "frmPlanPartOutFormula.frx":41D6
         FormatStyle(4)  =   "frmPlanPartOutFormula.frx":428A
         FormatStyle(5)  =   "frmPlanPartOutFormula.frx":4362
         ImageCount      =   0
         PrinterProperties=   "frmPlanPartOutFormula.frx":441A
      End
      Begin prjFarmManagement.uctlTextBox txtProgressPercent 
         Height          =   435
         Left            =   8700
         TabIndex        =   17
         Top             =   10080
         Width           =   1935
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin VB.Label lblProgressPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "lblProgressPercent"
         Height          =   315
         Left            =   7440
         TabIndex        =   18
         Top             =   10170
         Width           =   1245
      End
      Begin VB.Label lblFormulaNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFormulaNo"
         Height          =   315
         Left            =   9360
         TabIndex        =   15
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label lblPlanDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   480
         TabIndex        =   13
         Top             =   10110
         Width           =   1995
      End
      Begin Threed.SSCommand cmdSelect 
         Height          =   525
         Left            =   7080
         TabIndex        =   4
         Top             =   4440
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmPlanPartOutFormula.frx":45F2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   12960
         TabIndex        =   2
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmPlanPartOutFormula.frx":490C
         ButtonStyle     =   3
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   10
         Top             =   900
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   11280
         TabIndex        =   6
         Top             =   10020
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmPlanPartOutFormula.frx":4C26
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   12930
         TabIndex        =   7
         Top             =   10020
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmPlanPartOutFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Private m_Formula As CFormula
Private m_TempFormula As CFormula

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public TempCollection As Collection

Private m_TempCol1 As Collection
Private m_TempCol2 As Collection
Private m_TempCol3 As Collection
Private m_TempCol4 As Collection
'Private Sub PopulateDestColl()
'Dim Fl As CFormula
'
'   For Each Ri In TempCollection
'      Set D = New CFormula
'
'      If Ri.Flag <> "D" Then
'         D.BILLING_DOC_ID = Ri.DO_ID
'         D.DOCUMENT_DATE = Ri.DOCUMENT_DATE
'         D.DOCUMENT_NO = Ri.DOCUMENT_NO
'         D.TEMP_PAID_AMOUNT = Ri.PAID_AMOUNT
'         Call m_TempCol2.add(D)
'      End If
'
'      Set D = Nothing
'   Next Ri
'End Sub

Private Function IsIn(TempCol As Collection, TempID As Long) As Boolean
Dim D As CFormula
Dim Found As Boolean

   Found = False
   For Each D In TempCol
      If D.FORMULA_ID = TempID Then
         Found = True
      End If
   Next D
   
   IsIn = Found
End Function

Private Sub GenerateSourceItem(Rs As ADODB.Recordset, TempCol As Collection)
Dim Fl As CFormula
Dim X As Double

   Set m_TempCol1 = Nothing
   Set m_TempCol1 = New Collection
   While Not Rs.EOF
      Set Fl = New CFormula
      Call Fl.PopulateFromRS(1, Rs)
      
      If Not IsIn(m_TempCol2, Fl.FORMULA_ID) Then
            Call TempCol.add(Fl)
      End If
      
      Set Fl = Nothing
      Rs.MoveNext
   Wend
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
   
   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_Formula.FORMULA_NO = PatchWildCard(txtFormulaNo.Text)
      m_Formula.FROM_DATE = uctlFromDate.ShowDate
      m_Formula.TO_DATE = uctlToDate.ShowDate
      If Not glbProduction.QueryFormula(m_Formula, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
     End If
   End If
   
   If itemcount > 0 Then
      Call GenerateSourceItem(m_Rs, m_TempCol1)
      GridEX1.itemcount = m_TempCol1.Count
      GridEX1.Rebind
   Else
      GridEX1.itemcount = 0
      GridEX1.Rebind
   End If
   
   GridEX2.itemcount = m_TempCol2.Count
   GridEX2.Rebind
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
On Error GoTo ErrorHandler
Dim IsOK As Boolean
Dim Fi As CFormulaItem
Dim m_PlanPart As CPlanPart
Dim I As Long
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   If Not VerifyDate(lblPlanDate, uctlPlanDate, False) Then
      Exit Function
   End If
   
   Call glbDaily.StartTransaction
   
   Set m_PlanPart = New CPlanPart
   m_PlanPart.PLAN_DATE = uctlPlanDate.ShowDate
   m_PlanPart.PLAN_AREA = 1            'ประมาณการใช้
   Call m_PlanPart.UpdateCancelByDateArea
   Set m_PlanPart = Nothing
   
   For Each Fi In m_TempCol3
      Set m_PlanPart = New CPlanPart
      m_PlanPart.AddEditMode = SHOW_ADD
      m_PlanPart.PLAN_DATE = uctlPlanDate.ShowDate
      m_PlanPart.PLAN_AREA = 1            'ประมาณการใช้
      m_PlanPart.PART_ITEM_ID = Fi.PART_ITEM_ID
      
      m_PlanPart.PLAN_OUT = Fi.REAL_AMOUNT
         
      m_PlanPart.CANCEL_FLAG = "N"
   
      Call m_PlanPart.AddEditData
         
      I = I + 1
      txtProgressPercent.Text = MyDiff(I * 100, m_TempCol3.Count + m_TempCol4.Count)
      txtProgressPercent.Refresh
      
      Set m_PlanPart = Nothing
   Next Fi
   
   Set m_PlanPart = New CPlanPart
   m_PlanPart.PLAN_DATE = uctlPlanDate.ShowDate
   m_PlanPart.PLAN_AREA = 3            'ประมาณการผลิต
   Call m_PlanPart.UpdateCancelByDateArea
   Set m_PlanPart = Nothing
   
   For Each Fi In m_TempCol4
      Set m_PlanPart = New CPlanPart
      m_PlanPart.AddEditMode = SHOW_ADD
      m_PlanPart.PLAN_DATE = uctlPlanDate.ShowDate
      m_PlanPart.PLAN_AREA = 3            'ประมาณการผลิต
      m_PlanPart.PART_ITEM_ID = Fi.PART_ITEM_ID
      
      m_PlanPart.PLAN_IN = Fi.REAL_AMOUNT
         
      m_PlanPart.CANCEL_FLAG = "N"
   
      Call m_PlanPart.AddEditData
         
      I = I + 1
      txtProgressPercent.Text = MyDiff(I * 100, m_TempCol3.Count + m_TempCol4.Count)
      txtProgressPercent.Refresh
      
      Set m_PlanPart = Nothing
   Next Fi
   
   Call glbDaily.CommitTransaction
   
   Call EnableForm(Me, True)
   SaveData = True
   
   Exit Function
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
End Function
Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub
Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Public Sub CopyItem(TempCol1 As Collection, TempCol2 As Collection, ID As Long)
Dim L As CFormula
   
   If ID > 0 Then
      Set L = TempCol1(ID)
      
      frmPlanPartOutFormulaItem.HeaderText = "กรุณาใส่จำนวน BATCH"
      frmPlanPartOutFormulaItem.ShowMode = SHOW_EDIT
      frmPlanPartOutFormulaItem.ID = L.FORMULA_ID
      Set frmPlanPartOutFormulaItem.TempCollectionPlanOut = m_TempCol3
      Set frmPlanPartOutFormulaItem.TempCollectionPlanIn = m_TempCol4
      Load frmPlanPartOutFormulaItem
      frmPlanPartOutFormulaItem.Show 1
      
      OKClick = frmPlanPartOutFormulaItem.OKClick
      
      Unload frmPlanPartOutFormulaItem
      Set frmPlanPartOutFormulaItem = Nothing
      
      If OKClick Then
         L.Flag = "A"
         Call TempCol2.add(L)
         TempCol1.Remove (ID)
         
         GridEX3.itemcount = m_TempCol3.Count
         GridEX3.Rebind
         
         GridEX4.itemcount = m_TempCol4.Count
         GridEX4.Rebind
      End If
   End If
End Sub


Private Sub cmdSelect_Click()
Dim TempID As Long

   m_HasModify = True
   
   TempID = GridEX1.row
   Call CopyItem(m_TempCol1, m_TempCol2, TempID)

   GridEX1.itemcount = m_TempCol1.Count
   GridEX1.Rebind
   
   GridEX2.itemcount = m_TempCol2.Count
   GridEX2.Rebind
   
   GridEX3.itemcount = m_TempCol3.Count
   GridEX3.Rebind
   
   GridEX4.itemcount = m_TempCol4.Count
   GridEX4.Rebind
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      uctlPlanDate.ShowDate = Now
      
      Dim FirstDate As Date
      Dim LastDate As Date
      Call GetFirstLastDate(Now, FirstDate, LastDate)
      uctlFromDate.ShowDate = DateAdd("M", -1, FirstDate)
      uctlToDate.ShowDate = Now
      
      Call QueryData(True)
      
      Call EnableForm(Me, True)
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_Formula = Nothing
   Set m_TempCol1 = Nothing
   Set m_TempCol2 = Nothing
   Set m_TempCol3 = Nothing
   Set m_TempCol4 = Nothing
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 2085
   Col.Caption = "รหัสสูตร"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 2085
   Col.Caption = MapText("รายละเอียดสูตร")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 2475
   Col.Caption = MapText("วันที่สร้างสูตร")
 
   Set Col = GridEX1.Columns.add '7
   Col.Width = 3500
   Col.Caption = MapText("ผลิตภัณฑ์")
End Sub
Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX2.Columns.Clear
   GridEX2.BackColor = GLB_GRID_COLOR
   GridEX2.itemcount = 0
   GridEX2.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX2.ColumnHeaderFont.Bold = True
   GridEX2.ColumnHeaderFont.Name = GLB_FONT
   GridEX2.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX2.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX2.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX2.Columns.add '2
   Col.Width = 2085
   Col.Caption = "รหัสสูตร"

   Set Col = GridEX2.Columns.add '3
   Col.Width = 4320
   Col.Caption = MapText("รายละเอียดสูตร")
   
   Set Col = GridEX2.Columns.add '4
   Col.Width = 2475
   Col.Caption = MapText("วันที่สร้างสูตร")
 
   Set Col = GridEX2.Columns.add '7
   Col.Width = 3500
   Col.Caption = MapText("ผลิตภัณฑ์")
End Sub

Private Sub GetTotalPrice()
'Dim II As CExportItem
'Dim Sum As Double
'
'   Sum = 0
'   For Each II In m_Formula.ImportExports
'      If II.Flag <> "D" Then
'         Sum = Sum + CDbl(Format(II.EXPORT_AVG_PRICE, "0.00")) * CDbl(Format(II.EXPORT_AMOUNT, "0.00"))
'      End If
'   Next II
''
''   txtDeliveryFee.Text = Format(Sum, "0.00")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblFromDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblFormulaNo, MapText("รหัสสูตร"))
   Call InitNormalLabel(lblPlanDate, MapText("วันที่ประมาณการใช้"))
   
   Call InitNormalLabel(lblProgressPercent, MapText("% ความคืบหน้า"))
   
   txtProgressPercent.Enabled = False
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelect.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("คำนวณ (F2)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา"))
   Call InitMainButton(cmdSelect, MapText(">"))
   
   Call InitGrid1
   Call InitGrid2
   Call InitGrid3
   Call InitGrid4
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
   Set m_Formula = New CFormula
   Set m_TempCol1 = New Collection
   Set m_TempCol2 = New Collection
   Set m_TempCol3 = New Collection
   Set m_TempCol4 = New Collection
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim X As Double

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"


   If m_TempCol1 Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If
   
   Dim CR As CFormula
   If m_TempCol1.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol1, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.FORMULA_ID
   Values(2) = RealIndex
   Values(3) = CR.FORMULA_NO
   Values(4) = CR.FORMULA_DESC
   Values(5) = DateToStringExtEx2(CR.FORMULA_DATE)
   Values(6) = CR.PART_NO
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub GridEX2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"


   If m_TempCol2 Is Nothing Then
      Exit Sub
   End If
   
   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CFormula
   If m_TempCol2.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol2, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.FORMULA_ID
   Values(2) = RealIndex
   Values(3) = CR.FORMULA_NO
   Values(4) = CR.FORMULA_DESC
   Values(5) = DateToStringExtEx2(CR.FORMULA_DATE)
   Values(6) = CR.PART_NO
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub InitGrid3()
Dim Col As JSColumn

   GridEX3.Columns.Clear
   GridEX3.BackColor = GLB_GRID_COLOR
   GridEX3.itemcount = 0
   GridEX3.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX3.ColumnHeaderFont.Bold = True
   GridEX3.ColumnHeaderFont.Name = GLB_FONT
   GridEX3.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX3.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX3.Columns.add '2
   Col.Width = 1500
   Col.Caption = MapText("รหัสวัตถุดิบ")

   Set Col = GridEX3.Columns.add '3
   Col.Width = 3000
   Col.Caption = MapText("ชื่อวัตถุดิบ")
   
   Set Col = GridEX3.Columns.add '4
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("น้ำหนัก (Kg)")
End Sub
Private Sub InitGrid4()
Dim Col As JSColumn

   GridEX4.Columns.Clear
   GridEX4.BackColor = GLB_GRID_COLOR
   GridEX4.itemcount = 0
   GridEX4.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX4.ColumnHeaderFont.Bold = True
   GridEX4.ColumnHeaderFont.Name = GLB_FONT
   GridEX4.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX4.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX4.Columns.add '2
   Col.Width = 1500
   Col.Caption = MapText("รหัสผลิตภัณฑ์")

   Set Col = GridEX4.Columns.add '3
   Col.Width = 3000
   Col.Caption = MapText("ชื่อผลิตภัณฑ์")
   
   Set Col = GridEX4.Columns.add '4
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("น้ำหนัก (Kg)")
End Sub
Private Sub GridEX3_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_TempCol3 Is Nothing Then
      Exit Sub
   End If

     If RowIndex <= 0 Then
         Exit Sub
      End If

   Dim Ci As CFormulaItem
   If m_TempCol3.Count <= 0 Then
      Exit Sub
   End If
   Set Ci = GetItem(m_TempCol3, RowIndex, RealIndex)
   If Ci Is Nothing Then
      Exit Sub
   End If

   Values(1) = RealIndex
   Values(2) = Ci.PART_NO
   Values(3) = Ci.PART_ITEM_NAME
   Values(4) = FormatNumber(Ci.REAL_AMOUNT, 3)

   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub GridEX4_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_TempCol4 Is Nothing Then
      Exit Sub
   End If

     If RowIndex <= 0 Then
         Exit Sub
      End If

   Dim Ci As CFormulaItem
   If m_TempCol4.Count <= 0 Then
      Exit Sub
   End If
   Set Ci = GetItem(m_TempCol4, RowIndex, RealIndex)
   If Ci Is Nothing Then
      Exit Sub
   End If

   Values(1) = RealIndex
   Values(2) = Ci.PART_NO
   Values(3) = Ci.PART_ITEM_NAME
   Values(4) = FormatNumber(Ci.REAL_AMOUNT, 3)

   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

