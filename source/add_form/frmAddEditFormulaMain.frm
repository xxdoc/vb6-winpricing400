VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddEditFormulaMain 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditFormulaMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
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
      TabIndex        =   18
      Top             =   0
      Width           =   12720
      _ExtentX        =   22437
      _ExtentY        =   15690
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjFarmManagement.uctlTextLookup uctlApproveByLookup 
         Height          =   405
         Left            =   1590
         TabIndex        =   4
         Top             =   2760
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin VB.ComboBox cboFormulaType 
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2340
         Width           =   2625
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   10
         Top             =   4230
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
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2970
         Left            =   120
         TabIndex        =   11
         Top             =   4770
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   5239
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
         Column(1)       =   "frmAddEditFormulaMain.frx":27A2
         Column(2)       =   "frmAddEditFormulaMain.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditFormulaMain.frx":290E
         FormatStyle(2)  =   "frmAddEditFormulaMain.frx":2A6A
         FormatStyle(3)  =   "frmAddEditFormulaMain.frx":2B1A
         FormatStyle(4)  =   "frmAddEditFormulaMain.frx":2BCE
         FormatStyle(5)  =   "frmAddEditFormulaMain.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditFormulaMain.frx":2D5E
      End
      Begin prjFarmManagement.uctlDate uctlFormulaDate 
         Height          =   405
         Left            =   1590
         TabIndex        =   2
         Top             =   1890
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtFormulaNo 
         Height          =   435
         Left            =   1590
         TabIndex        =   0
         Top             =   990
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtFormulaDesc 
         Height          =   435
         Left            =   1590
         TabIndex        =   1
         Top             =   1440
         Width           =   5985
         _ExtentX        =   6535
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlPartItemLookup 
         Height          =   405
         Left            =   1590
         TabIndex        =   5
         Top             =   3180
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtTotalRatio 
         Height          =   435
         Left            =   8490
         TabIndex        =   9
         Top             =   1950
         Width           =   1695
         _ExtentX        =   4736
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlLocationLookup 
         Height          =   405
         Left            =   1590
         TabIndex        =   6
         Top             =   3600
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtRMC 
         Height          =   435
         Left            =   8490
         TabIndex        =   29
         Top             =   2400
         Width           =   1695
         _ExtentX        =   4736
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPMC 
         Height          =   435
         Left            =   8490
         TabIndex        =   32
         Top             =   2850
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   4736
         _ExtentY        =   767
      End
      Begin Threed.SSCheck chkCancelFlag 
         Height          =   375
         Left            =   8490
         TabIndex        =   36
         Top             =   1080
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCommand cmdOther 
         Height          =   525
         Left            =   5160
         TabIndex        =   35
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditFormulaMain.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdCalculate 
         Height          =   525
         Left            =   6870
         TabIndex        =   15
         Top             =   7830
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditFormulaMain.frx":3250
         ButtonStyle     =   3
      End
      Begin VB.Label lblPMC 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFormulaNo"
         Height          =   315
         Left            =   7110
         TabIndex        =   34
         Top             =   2940
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label4 
         Caption         =   "lblFormulaNo"
         Height          =   315
         Left            =   10230
         TabIndex        =   33
         Top             =   2970
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblRMC 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFormulaNo"
         Height          =   315
         Left            =   7110
         TabIndex        =   31
         Top             =   2490
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "lblFormulaNo"
         Height          =   315
         Left            =   10230
         TabIndex        =   30
         Top             =   2520
         Width           =   1305
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFormulaItem"
         Height          =   315
         Left            =   150
         TabIndex        =   28
         Top             =   3720
         Width           =   1365
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   8490
         TabIndex        =   7
         Top             =   3300
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditFormulaMain.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   10110
         TabIndex        =   8
         Top             =   3300
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label Label2 
         Caption         =   "lblFormulaNo"
         Height          =   315
         Left            =   10230
         TabIndex        =   27
         Top             =   2070
         Width           =   1305
      End
      Begin VB.Label lblTotalRatio 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFormulaNo"
         Height          =   315
         Left            =   7110
         TabIndex        =   26
         Top             =   2040
         Width           =   1305
      End
      Begin VB.Label lblFormulaNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFormulaNo"
         Height          =   315
         Left            =   210
         TabIndex        =   25
         Top             =   1110
         Width           =   1305
      End
      Begin VB.Label lblFormulaItem 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFormulaItem"
         Height          =   315
         Left            =   150
         TabIndex        =   24
         Top             =   3300
         Width           =   1365
      End
      Begin VB.Label lblFormulaDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFormulaDesc"
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   1530
         Width           =   1395
      End
      Begin VB.Label lblFormulaType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFormulaType"
         Height          =   315
         Left            =   330
         TabIndex        =   22
         Top             =   2430
         Width           =   1185
      End
      Begin VB.Label lblFormulaApp 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFormulaApp"
         Height          =   315
         Left            =   180
         TabIndex        =   21
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblFormulaDate 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFormulaDate"
         Height          =   315
         Left            =   210
         TabIndex        =   20
         Top             =   2010
         Width           =   1305
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8520
         TabIndex        =   16
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditFormulaMain.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10170
         TabIndex        =   17
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   13
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   12
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditFormulaMain.frx":3B9E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   14
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditFormulaMain"
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
Private m_Formulas As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long

Private FileName As String
Private m_SumUnit As Double
Private m_PartTypes  As Collection
Private m_PartItems As Collection
Private m_Locations As Collection

Public TempCollection As Collection

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_Formula.FORMULA_ID = id
      m_Formula.QueryFlag = 1
      If Not glbProduction.QueryFormula(m_Formula, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
        Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Formula.PopulateFromRS(1, m_Rs)

      txtFormulaNo.Text = m_Formula.FORMULA_NO
      txtFormulaDesc.Text = m_Formula.FORMULA_DESC
      uctlFormulaDate.ShowDate = m_Formula.FORMULA_DATE
      uctlApproveByLookup.MyCombo.ListIndex = IDToListIndex(uctlApproveByLookup.MyCombo, m_Formula.PART_TYPE_ID)
      uctlPartItemLookup.MyCombo.ListIndex = IDToListIndex(uctlPartItemLookup.MyCombo, m_Formula.PART_ITEM_ID)
      cboFormulaType.ListIndex = IDToListIndex(cboFormulaType, m_Formula.FORMULA_TYPE)
      uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, m_Formula.LOCATION_ID)
      chkCancelFlag.Value = FlagToCheck(m_Formula.CANCEL_FLAG)
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call TabStrip1_Click
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Per As Double
Dim Ac As CFormulaItem
   If ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("PRODUCT_FORMULA_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If
   If Not VerifyTextControl(lblFormulaNo, txtFormulaNo, False) Then
      Exit Function
  End If
   
   If Not VerifyTextControl(lblFormulaDesc, txtFormulaDesc, False) Then
      Exit Function
   End If
   
   If Not VerifyDate(lblFormulaDate, uctlFormulaDate, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblFormulaApp, uctlApproveByLookup.MyCombo, True) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblFormulaType, cboFormulaType, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblFormulaItem, uctlPartItemLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(FORMULA_NO, txtFormulaNo.Text, id) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtFormulaNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
'   Per = 0
'   For Each Ac In m_Formula.Inputs
'   If Ac.Flag <> "D" Then
'      Per = Per + Ac.ITEM_PERCENT
'   End If
'   Next Ac
'
'   If Per <> 100 Then
'      glbErrorLog.LocalErrorMsg = "จำนวนเปอร์เซนต์วัตถุดิบไม่เท่ากับ 100 เปอร์เซ็นต์"
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Formula.FORMULA_ID = id
   m_Formula.AddEditMode = ShowMode
   m_Formula.FORMULA_NO = txtFormulaNo.Text
   m_Formula.FORMULA_DESC = txtFormulaDesc.Text
   m_Formula.FORMULA_DATE = uctlFormulaDate.ShowDate
   m_Formula.APPROVED_BY = -1
   m_Formula.FORMULA_TYPE = cboFormulaType.ItemData(Minus2Zero(cboFormulaType.ListIndex))
   m_Formula.PART_ITEM_ID = uctlPartItemLookup.MyCombo.ItemData(Minus2Zero(uctlPartItemLookup.MyCombo.ListIndex))
   m_Formula.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   m_Formula.PMC = Val(txtPMC.Text)
   m_Formula.RMC = Val(Replace(txtRMC.Text, ",", ""))
   m_Formula.CANCEL_FLAG = Check2Flag(chkCancelFlag.Value)
   
   Call EnableForm(Me, False)
   If Not glbProduction.AddEditFormula(m_Formula, IsOK, True, glbErrorLog) Then
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
Private Sub cboFormulaApp_Change()
m_HasModify = True
End Sub

Private Sub cboFormulaApp_Click()
m_HasModify = True
End Sub

Private Sub cboFormulaItem_Click()
m_HasModify = True
End Sub

Private Sub cboFormulaType_Click()
   m_HasModify = True
End Sub

Private Sub cboFormulaType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub CalculateTotalRatio()
Dim D As CFormulaItem
Dim Sum As Double
Dim Sum2 As Double
Dim Fv As CFormulaVariable
Dim Loss As Double
Dim Can As Double
Dim OverHead As Double
Dim Pmc1 As Double
Dim Rmc1 As Double
Dim Sum3 As Double

   Sum = 0
   Sum2 = 0
   Sum3 = 0
   For Each D In m_Formula.Inputs
      If D.Flag <> "D" Then
         Sum = Sum + D.ITEM_PERCENT
         Sum3 = Sum3 + D.REAL_AMOUNT
         Sum2 = Sum2 + D.ITEM_PERCENT * D.AVG_PRICE
      End If
   Next D
   
   If Not (m_Formula.FormulaVariables Is Nothing) Then
   If m_Formula.FormulaVariables.Count > 0 Then
      Set Fv = m_Formula.FormulaVariables("1")
      Loss = Fv.VARIABLE_VALUE
      Set Fv = m_Formula.FormulaVariables("2")
      OverHead = Fv.VARIABLE_VALUE
      Set Fv = m_Formula.FormulaVariables("3")
      Can = Fv.VARIABLE_VALUE
      Call glbProduction.CalPMC2(Loss, Can, OverHead, Sum2, Rmc1, Pmc1)
      txtRMC.Text = FormatNumber(Rmc1)
      txtPMC.Text = FormatNumber(Pmc1)
      End If
   End If
   
   txtTotalRatio.Text = FormatNumber(Sum, 3)
   txtRMC.Text = FormatNumber(Sum3, 3)
End Sub

Private Sub chkCancelFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If

   OKClick = False
    If TabStrip1.SelectedItem.Index = 1 Then
     Set frmAddEditFormulaInput.TempCollection = m_Formula.Inputs
      frmAddEditFormulaInput.ParentShowMode = ShowMode
      frmAddEditFormulaInput.ShowMode = SHOW_ADD
      frmAddEditFormulaInput.HeaderText = MapText("เพิ่มอัตราส่วนวัตถุดิบ")
      Load frmAddEditFormulaInput
      frmAddEditFormulaInput.Show 1

      OKClick = frmAddEditFormulaInput.OKClick

      Unload frmAddEditFormulaInput
      Set frmAddEditFormulaInput = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Formula.Inputs)
         GridEX1.Rebind
         Call CalculateTotalRatio
      End If
   End If

   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdCalculate_Click()
Dim D As CFormulaItem
Dim M As CFormulaVariable
Dim Sum As Double
Dim SumMarkup As Double
Dim FifoPrice As Double
Dim IsOK As Boolean
Dim LastPMC As Double
Dim AvgPMC As Double
Dim PL As CPartLocation
Dim TempRs As ADODB.Recordset
Dim iCount As Long

   Call EnableForm(Me, False)
   
   SumMarkup = 0
   Sum = 0

   'ใช้ราคาเฉลี่ยเท่านั้นเพราะไม่รู้ปริมาณจริงที่จะไปตัดสต็อค
   Set m_Formulas = Nothing
   Set m_Formulas = New Collection
   For Each D In m_Formula.Inputs
      If D.FROM_FORMULA > 0 Then
         Call glbProduction.GetCalculatedPrice(D.FROM_FORMULA, AvgPMC, 1, D.ITEM_PERCENT, IsOK, glbErrorLog)
      Else
         Set TempRs = New ADODB.Recordset
         Set PL = New CPartLocation
         PL.PART_LOCATION_ID = -1
         PL.PART_ITEM_ID = D.PART_ITEM_ID
         PL.LOCATION_ID = D.LOCATION_ID
         Call PL.QueryData(1, TempRs, iCount)
         If Not TempRs.EOF Then
            Call PL.PopulateFromRS(TempRs)
         End If
         AvgPMC = PL.AVG_PRICE
         Set PL = Nothing
         If TempRs.State = adStateOpen Then
            TempRs.Close
         End If
         Set TempRs = Nothing
      End If

      If D.Flag <> "D" Then
         D.AVG_PRICE = AvgPMC
         If D.Flag <> "A" Then
            D.Flag = "E"
         End If
      End If
   Next D
   
   Call TabStrip1_Click
   
   m_HasModify = True
   Call EnableForm(Me, True)
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
         m_Formula.Inputs.Remove (ID2)
      Else
         m_Formula.Inputs.Item(ID2).Flag = "D"
      End If

      Call ReArrangeRatio(m_Formula.Inputs)
      
      GridEX1.ItemCount = CountItem(m_Formula.Inputs)
      GridEX1.Rebind
      Call CalculateTotalRatio
      m_HasModify = True
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
Dim OKClick As Boolean
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   id = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
     Set frmAddEditFormulaInput.TempCollection = m_Formula.Inputs
      frmAddEditFormulaInput.id = id
      frmAddEditFormulaInput.ShowMode = SHOW_EDIT
      frmAddEditFormulaInput.HeaderText = MapText("แก้ไขอัตราส่วนวัตถุดิบ")
      Load frmAddEditFormulaInput
      frmAddEditFormulaInput.Show 1

      OKClick = frmAddEditFormulaInput.OKClick

      Unload frmAddEditFormulaInput
      Set frmAddEditFormulaInput = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Formula.Inputs)
         GridEX1.Rebind
         Call CalculateTotalRatio
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
     Set frmEditVariable.TempCollection = m_Formula.FormulaVariables
      frmEditVariable.id = id
      frmEditVariable.ShowMode = SHOW_EDIT
      frmEditVariable.HeaderText = MapText("แก้ไขตัวแปรคำนวณ")
      Load frmEditVariable
      frmEditVariable.Show 1

      OKClick = frmEditVariable.OKClick

      Unload frmEditVariable
      Set frmEditVariable = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Formula.FormulaVariables)
         GridEX1.Rebind
      End If
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdOK_Click()
Dim Ac As CFormulaItem
Dim Per As Double
 
   If cmdOK.Enabled = False Then
      Exit Sub
   End If

   If Not SaveData Then
      Exit Sub
   End If
      
   OKClick = True
   Unload Me
End Sub


Private Sub cmdOther_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("Refresh ราคาเฉลี่ย")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   If lMenuChosen = 1 Then
        Call RefreshAvgPrice
        m_HasModify = True
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

   If m_HasModify Or ((Not m_HasModify) And (ShowMode = SHOW_ADD)) Then
      glbErrorLog.LocalErrorMsg = "กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("สูตร 100", "ปรับค่าหน้ากระดาษ", "-", "สูตร 100 เหมือนจริง", "ปรับค่าหน้ากระดาษ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   
   If lMenuChosen = 1 Then
      ReportKey = "CReportFormula001"
      
      Set Report = New CReportFormula001
      ReportFlag = True
   ElseIf lMenuChosen = 2 Then
      ReportKey = "CReportFormula001"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("สูตรการผลิต")
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 4 Then
      ReportKey = "CReportFormula002"

      Set Report = New CReportFormula002
      ReportFlag = True
   ElseIf lMenuChosen = 5 Then
      ReportKey = "CReportFormula002"

      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("สูตรการผลิต")
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   End If
   
   If Not Report Is Nothing Then
      Call Report.AddParam(m_Formula.FORMULA_ID, "FORMULA_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
   End If
   
   If ReportFlag Then
      Set frmReport.ReportObject = Report
      frmReport.HeaderText = pnlHeader.Caption
      Load frmReport
      frmReport.Show 1
   
      Unload frmReport
      Set frmReport = Nothing
      Set Report = Nothing
   Else
      frmReportConfig.ReportMode = 1
      frmReportConfig.ShowMode = EditMode
      frmReportConfig.id = Rc.REPORT_CONFIG_ID
      frmReportConfig.ReportKey = ReportKey
      frmReportConfig.HeaderText = HeaderText
      Load frmReportConfig
      frmReportConfig.Show 1
      
      Unload frmReportConfig
      Set frmReportConfig = Nothing
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdSave_Click()
Dim Result As Boolean
   If Not SaveData Then
      Exit Sub
   End If
   
   ShowMode = SHOW_EDIT
   id = m_Formula.FORMULA_ID
   m_Formula.QueryFlag = 1
   QueryData (True)
   m_HasModify = False
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadFormulaType(cboFormulaType)
      
      Call LoadPartType(uctlApproveByLookup.MyCombo, m_PartTypes)
      Set uctlApproveByLookup.MyCollection = m_PartTypes
      
'      Call LoadPartItem(uctlPartItemLookup.MyCombo, m_PartItems)
'      Set uctlPartItemLookup.MyCollection = m_PartItems
      
      Call LoadLocation(uctlLocationLookup.MyCombo, m_Locations, 2)
      Set uctlLocationLookup.MyCollection = m_Locations
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Formula.QueryFlag = 1
         Call QueryData(True)
         
      ElseIf ShowMode = SHOW_ADD Then
         uctlFormulaDate.ShowDate = Now
         Call LoadJobVariableEx(Nothing)
         
        m_Formula.QueryFlag = 0
         Call QueryData(False)
      End If
      
'      Call TabStrip1_Click
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub

Public Sub LoadJobVariableEx(C As ComboBox)
On Error GoTo ErrorHandler
Dim D As CFormulaVariable
Dim ItemCount As Long
Dim I As Long
Dim TempData As CFormulaVariable

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (m_Formula.FormulaVariables Is Nothing) Then
      Set m_Formula.FormulaVariables = Nothing
      Set m_Formula.FormulaVariables = New Collection
   End If
   
   For I = 1 To 3
      Set TempData = New CFormulaVariable
      TempData.VARIABLE_ID = I
      TempData.VARIABLE_NAME = VariableToText(I)
      TempData.Flag = "A"
      
      If Not (C Is Nothing) Then
         C.AddItem (TempData.VARIABLE_NAME)
         C.ItemData(I) = TempData.VARIABLE_ID
      End If
   
      If Not (m_Formula.FormulaVariables Is Nothing) Then
         Call m_Formula.FormulaVariables.add(TempData, Trim(str(TempData.VARIABLE_ID)))
      End If
      Set TempData = Nothing
   Next I
      
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
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
   
   Set m_Formula = Nothing
   Set m_Formulas = Nothing
   Set m_PartTypes = Nothing
   Set m_PartItems = Nothing
   Set m_Locations = Nothing
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
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"
   Col.Visible = False
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   Col.Visible = False
   
   GridEX1.Columns.Item(2).Visible = False
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2100
   Col.Caption = MapText("รหัสวัตถุดิบ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 4815
   Col.Caption = MapText("ชื่อวัตถุดิบ")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2580
   Col.Caption = MapText("ประเภทวัตถุดิบ")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 1845
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("น้ำหนัก (Kg)")

   Set Col = GridEX1.Columns.add '7
   Col.Width = 1845
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("อัตราส่วน (%)")
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
   Col.Visible = False
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   Col.Visible = False

   Set Col = GridEX1.Columns.add '3
   Col.Width = 9450
   Col.Caption = MapText("ตัวแปร")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2145
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวน")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblFormulaNo, MapText("รหัสสูตร"))
   Call InitNormalLabel(lblFormulaDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblFormulaDate, MapText("วันที่สร้างสูตร"))
   Call InitNormalLabel(lblFormulaApp, MapText("ประเภท"))
   Call InitNormalLabel(lblFormulaType, MapText("ประเภทสูตร"))
   Call InitNormalLabel(lblFormulaItem, MapText("ผลิตภัณฑ์"))
   Call InitNormalLabel(lblTotalRatio, MapText("อัตราส่วนรวม"))
   Call InitNormalLabel(Label2, MapText("%"))
   Call InitNormalLabel(lblLocation, MapText("สถานที่จัดเก็บ"))
   Call InitNormalLabel(lblRMC, MapText("น้ำหนักรวม"))
   Call InitNormalLabel(lblPMC, MapText("PMC"))
   Call InitNormalLabel(Label1, MapText("Kg"))
   Call InitNormalLabel(Label4, MapText(""))
   
   Call txtFormulaNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtFormulaDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtTotalRatio.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalRatio.Enabled = False
   Call txtRMC.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtRMC.Enabled = False
   Call txtPMC.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtPMC.Enabled = False
   
   Call InitCombo(cboFormulaType)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitGrid1
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdCalculate.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOther.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSave, MapText("บันทึก"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdCalculate, MapText("คำนวณ"))
   Call InitMainButton(cmdOther, MapText("อื่นๆ"))
   
   Call InitCheckBox(chkCancelFlag, "ยกเลิกใช้งาน")
   
   If ShowMode = SHOW_VIEW_ONLY Then
      cmdOK.Enabled = False
      cmdAdd.Enabled = False
      cmdDelete.Enabled = False
      cmdSave.Enabled = False
   End If
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("วัตถุดิบที่ใช้")
'   TabStrip1.Tabs.add().Caption = MapText("ตัวแปรคำนวณ")
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
   Set m_Formulas = New Collection
   Set m_PartTypes = New Collection
   Set TempCollection = New Collection
   Set m_PartItems = New Collection
   Set m_Locations = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

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
     If m_Formula.Inputs Is Nothing Then
         Exit Sub
      End If

     If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Ci As CFormulaItem
      If m_Formula.Inputs.Count <= 0 Then
         Exit Sub
      End If
      Set Ci = GetItem(m_Formula.Inputs, RowIndex, RealIndex)
      If Ci Is Nothing Then
         Exit Sub
      End If
      Values(1) = Ci.FORMULA_ITEM_ID
      Values(2) = RealIndex
      Values(3) = Ci.PART_NO
      Values(4) = Ci.PART_ITEM_NAME
      Values(5) = Ci.PART_TYPE_NAME
      Values(6) = FormatNumber(Ci.REAL_AMOUNT, 3)
      Values(7) = FormatNumber(Ci.ITEM_PERCENT, 3)
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
     If m_Formula.FormulaVariables Is Nothing Then
         Exit Sub
      End If

     If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Jv As CFormulaVariable
      If m_Formula.FormulaVariables.Count <= 0 Then
         Exit Sub
      End If
      Set Jv = GetItem(m_Formula.FormulaVariables, RowIndex, RealIndex)
      If Jv Is Nothing Then
         Exit Sub
      End If
      Values(1) = Jv.VARIABLE_ID
      Values(2) = RealIndex
      Values(3) = VariableToText(Jv.VARIABLE_ID)
      Values(4) = FormatNumber(Jv.VARIABLE_VALUE)
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
      cmdAdd.Enabled = Not (ShowMode = SHOW_VIEW_ONLY)
      cmdEdit.Enabled = True
      cmdDelete.Enabled = Not (ShowMode = SHOW_VIEW_ONLY)
      
      Call InitGrid1
      GridEX1.ItemCount = CountItem(m_Formula.Inputs)
      GridEX1.Rebind
      Call CalculateTotalRatio
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      cmdAdd.Enabled = False
      cmdEdit.Enabled = True
      cmdDelete.Enabled = False
      
      Call InitGrid2
      GridEX1.ItemCount = CountItem(m_Formula.FormulaVariables)
      GridEX1.Rebind
   End If
End Sub

Private Sub txtFormulaDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtFormulaNo_Change()
   m_HasModify = True
End Sub



Private Sub uctlApproveByLookup_Change()
Dim PartTypeID As Long
Dim Pt As CPartType

   PartTypeID = uctlApproveByLookup.MyCombo.ItemData(Minus2Zero(uctlApproveByLookup.MyCombo.ListIndex))
   
   If PartTypeID > 0 Then
      Set Pt = GetPartType(m_PartTypes, Trim(str(PartTypeID)))
      Call LoadPartItem(uctlPartItemLookup.MyCombo, m_PartItems, PartTypeID, "N")
      Set uctlPartItemLookup.MyCollection = m_PartItems
   
      Call LoadLocation(uctlLocationLookup.MyCombo, m_Locations, 2, , , Pt.PART_GROUP_ID)
      Set uctlLocationLookup.MyCollection = m_Locations
   End If
   
   m_HasModify = True
End Sub

Private Sub uctlFormulaDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartItemLookup_Change()
   m_HasModify = True
End Sub
Private Sub RefreshAvgPrice()
Dim Ba As CBalanceAccum
Dim TempColl As Collection
Dim FrItem As CFormulaItem
    Set TempColl = New Collection
    Call LoadInventoryPartBalance(Nothing, TempColl, uctlFormulaDate.ShowDate)
    
    For Each FrItem In m_Formula.Inputs
        Set Ba = GetBalanceAccum(TempColl, Trim(str(FrItem.PART_ITEM_ID)))
        FrItem.Flag = "E"
        FrItem.AVG_PRICE = Ba.AVG_PRICE
    Next FrItem
    Set Ba = Nothing
    Set TempColl = Nothing
End Sub
