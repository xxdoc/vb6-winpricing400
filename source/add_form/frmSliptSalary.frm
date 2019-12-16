VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmSliptSalary 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmSliptSalary.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboName 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   960
      Width           =   2955
   End
   Begin VB.ComboBox cboMonth 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   1440
      Width           =   2955
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   375
         Left            =   180
         TabIndex        =   23
         Top             =   2400
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cboPosition 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   960
         Width           =   2955
      End
      Begin VB.ComboBox cboYear 
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1440
         Width           =   2955
      End
      Begin prjFarmManagement.uctlTextBox txtSalaryAdd 
         Height          =   435
         Left            =   8760
         TabIndex        =   2
         Top             =   3720
         Width           =   2145
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5055
         Left            =   180
         TabIndex        =   6
         Top             =   2640
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   8916
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
         Column(1)       =   "frmSliptSalary.frx":27A2
         Column(2)       =   "frmSliptSalary.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmSliptSalary.frx":290E
         FormatStyle(2)  =   "frmSliptSalary.frx":2A6A
         FormatStyle(3)  =   "frmSliptSalary.frx":2B1A
         FormatStyle(4)  =   "frmSliptSalary.frx":2BCE
         FormatStyle(5)  =   "frmSliptSalary.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmSliptSalary.frx":2D5E
      End
      Begin prjFarmManagement.uctlTextBox txtSalary 
         Height          =   435
         Left            =   8760
         TabIndex        =   0
         Top             =   2400
         Width           =   2145
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLendReMinder 
         Height          =   435
         Left            =   8760
         TabIndex        =   3
         Top             =   4920
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtSalarySub 
         Height          =   435
         Left            =   8760
         TabIndex        =   24
         Top             =   4320
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotSalary 
         Height          =   435
         Left            =   8760
         TabIndex        =   26
         Top             =   6720
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtoldSalary 
         Height          =   435
         Left            =   8760
         TabIndex        =   32
         Top             =   3240
         Width           =   2145
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdDeleteMain 
         Height          =   525
         Left            =   5280
         TabIndex        =   35
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmSliptSalary.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblCurrentSalary 
         Alignment       =   1  'Right Justify
         Caption         =   "lblCurrentSalary"
         Height          =   315
         Left            =   6960
         TabIndex        =   34
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label lblBath0 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBath0"
         Height          =   315
         Left            =   10920
         TabIndex        =   33
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblBath5 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBath5"
         Height          =   255
         Left            =   11040
         TabIndex        =   31
         Top             =   6840
         Width           =   495
      End
      Begin VB.Label lblBath1 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBath1"
         Height          =   315
         Left            =   10920
         TabIndex        =   30
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label lblBath2 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBath2"
         Height          =   315
         Left            =   11040
         TabIndex        =   29
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label lblBath3 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBath3"
         Height          =   315
         Left            =   11040
         TabIndex        =   28
         Top             =   4440
         Width           =   495
      End
      Begin VB.Label lblBath4 
         Alignment       =   1  'Right Justify
         Caption         =   "lblBath4"
         Height          =   315
         Left            =   11040
         TabIndex        =   27
         Top             =   5040
         Width           =   495
      End
      Begin VB.Label lblTotSalary 
         Alignment       =   1  'Right Justify
         Caption         =   "lblTotSalary"
         Height          =   315
         Left            =   6840
         TabIndex        =   25
         Top             =   6840
         Width           =   1815
      End
      Begin VB.Label lblSalarySub 
         Alignment       =   1  'Right Justify
         Caption         =   "lblSalarySub"
         Height          =   435
         Left            =   6960
         TabIndex        =   20
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label lblMonth 
         Alignment       =   1  'Right Justify
         Caption         =   "lblMonth"
         Height          =   315
         Left            =   0
         TabIndex        =   19
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblLendReMinder 
         Alignment       =   1  'Right Justify
         Caption         =   "lblLendReMinder"
         Height          =   435
         Left            =   6840
         TabIndex        =   18
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label lbloldSalary 
         Alignment       =   1  'Right Justify
         Caption         =   "lbloldSalary"
         Height          =   315
         Left            =   6960
         TabIndex        =   17
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label lblPosition 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPosition"
         Height          =   315
         Left            =   0
         TabIndex        =   16
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblSalaryAdd 
         Alignment       =   1  'Right Justify
         Caption         =   "lblSalaryAdd"
         Height          =   435
         Left            =   6960
         TabIndex        =   15
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label lblYear 
         Alignment       =   1  'Right Justify
         Caption         =   "lblMonth"
         Height          =   315
         Left            =   4560
         TabIndex        =   14
         Top             =   1560
         Width           =   615
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   9000
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3390
         TabIndex        =   9
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmSliptSalary.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   7
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmSliptSalary.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   8
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10095
         TabIndex        =   11
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8445
         TabIndex        =   10
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmSliptSalary.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmSliptSalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_Employee As CEmployee
Private m_TempEmployee As CEmployee
Private m_Rs As ADODB.Recordset
Private m_Rs1 As ADODB.Recordset
Private m_HasModify As Boolean
Private m_TableName As String
Public ShowMode As SHOW_MODE_TYPE
Private m_sliptSalary As CSliptSalary
Private m_TempSliptSalary As CSliptSalary
Private m_EmpReceivable As CEmpReceivable
Public OKClick As Boolean
Public TempRemind As Double
Public id As Long

Private Sub cboMonth_Click()
Dim id As Long
id = cboMonth.ItemData(Minus2Zero(cboMonth.ListIndex))

If id <> 0 Then
          cboYear.Enabled = True
         cboMonth.Enabled = False
Call QueryData(True)
End If
End Sub

Private Sub cboName_Change()
Call QueryData(True)
End Sub

Private Sub cboName_Click()
Dim id As Long
id = cboName.ItemData(Minus2Zero(cboName.ListIndex))
If id <> 0 Then
         cboMonth.Enabled = True
         cboName.Enabled = False
         Call QueryData(True)
         End If
End Sub

Private Sub cboPosition_Click()
Dim id As Long

   id = cboPosition.ItemData(Minus2Zero(cboPosition.ListIndex))
   If id <> 0 Then
   Call LoadEmployee(cboName, , id)
   cboName.Enabled = True
   cboPosition.Enabled = False
   End If
End Sub


Private Sub cboYear_Click()
Dim id As Long
id = cboYear.ListIndex
If id <> 0 Then
  cboYear.Enabled = False
         cmdAdd.Enabled = True
       cmdEdit.Enabled = True
       cmdDelete.Enabled = True
Call QueryData(True)
End If
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
      
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditSliptSalary.TempCollection = m_sliptSalary.SliptAdd
      frmAddEditSliptSalary.ParentShowMode = ShowMode
      frmAddEditSliptSalary.ShowMode = SHOW_ADD
      frmAddEditSliptSalary.HeaderText = MapText("เพิ่มส่วนบวกเงินเดือน")
      Load frmAddEditSliptSalary
      frmAddEditSliptSalary.Show 1

      OKClick = frmAddEditSliptSalary.OKClick
      Unload frmAddEditSliptSalary
      Set frmAddEditSliptSalary = Nothing
         GridEX1.ItemCount = CountItem(m_sliptSalary.SliptAdd)
        GridEX1.Rebind
       Call CalculateSliptAdd
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      
      Set frmAddEditSliptSalarySub.TempCollection = m_sliptSalary.SliptSub
      frmAddEditSliptSalarySub.ParentShowMode = ShowMode
      frmAddEditSliptSalarySub.ShowMode = SHOW_ADD
      frmAddEditSliptSalarySub.HeaderText = MapText("เพิ่มส่วนหักเงินเดือน")
      Load frmAddEditSliptSalarySub
      frmAddEditSliptSalarySub.Show 1

      OKClick = frmAddEditSliptSalarySub.OKClick
      Unload frmAddEditSliptSalarySub
      Set frmAddEditSliptSalarySub = Nothing
         GridEX1.ItemCount = CountItem(m_sliptSalary.SliptSub)
        GridEX1.Rebind
       Call CalculateSliptAdd

End If
m_HasModify = True
End Sub

Private Sub cmdClear_Click()
   cboPosition.ListIndex = -1
   cboName.ListIndex = -1
   cboMonth.ListIndex = -1
   cboYear.ListIndex = -1
   cboName.Enabled = False
   cboPosition.Enabled = True
   cboMonth.Enabled = False
   cboYear.Enabled = False
      txtSalary.Text = ""
   txtSalaryAdd.Text = ""
   txtSalarySub.Text = ""
   txtLendReMinder.Text = ""
   txtTotSalary.Text = ""
   txtoldSalary.Text = ""
Set m_sliptSalary.SliptAdd = Nothing
Set m_sliptSalary.SliptAdd = New Collection
Set m_sliptSalary.SliptSub = Nothing
Set m_sliptSalary.SliptSub = New Collection
TempRemind = 0
m_HasModify = False
Call QueryData(True)
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
         m_sliptSalary.SliptAdd.Remove (ID2)
      Else
         m_sliptSalary.SliptAdd.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_sliptSalary.SliptAdd)
      GridEX1.Rebind
      m_HasModify = True
      Call CalculateSliptAdd
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If ID1 <= 0 Then
         m_sliptSalary.SliptSub.Remove (ID2)
      Else
         m_sliptSalary.SliptSub.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_sliptSalary.SliptSub)
      GridEX1.Rebind
      m_HasModify = True
   Call CalculateSliptAdd
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
     Set frmAddEditSliptSalary.TempCollection = m_sliptSalary.SliptAdd
      frmAddEditSliptSalary.id = id
      frmAddEditSliptSalary.ShowMode = SHOW_EDIT
      frmAddEditSliptSalary.HeaderText = MapText("แก้ไขส่วนเพิ่มเงินเดือน")
      Load frmAddEditSliptSalary
      frmAddEditSliptSalary.Show 1

      OKClick = frmAddEditSliptSalary.OKClick

      Unload frmAddEditSliptSalary
      Set frmAddEditSliptSalary = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_sliptSalary.SliptAdd)
         GridEX1.Rebind
      End If
       Call CalculateSliptAdd
   
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
     Set frmAddEditSliptSalarySub.TempCollection = m_sliptSalary.SliptSub
      frmAddEditSliptSalarySub.id = id
      frmAddEditSliptSalarySub.ShowMode = SHOW_EDIT
      frmAddEditSliptSalarySub.HeaderText = MapText("แก้ไขส่วนหักเงินเดือน")
      Load frmAddEditSliptSalarySub
      frmAddEditSliptSalarySub.Show 1

      OKClick = frmAddEditSliptSalarySub.OKClick

      Unload frmAddEditSliptSalarySub
      Set frmAddEditSliptSalarySub = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_sliptSalary.SliptSub)
         GridEX1.Rebind
      End If
End If
       Call CalculateSliptAdd

End Sub

Private Sub cmdOK_Click()
      If Not SaveData Then
      Exit Sub
   End If
   OKClick = True
   Unload Me
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call LoadPosition(cboPosition)
      Call LoadEmployee(cboName, , -1)
      cboName.Enabled = False
      cboMonth.Enabled = False
      cboYear.Enabled = False
      Call QueryData(True)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim ItemCount1 As Long
Dim Temp As Long
    cmdAdd.Enabled = False
       cmdEdit.Enabled = False
       cmdDelete.Enabled = False
   cmdDeleteMain.Enabled = False
   If Flag Then
      Call EnableForm(Me, False)
            
      m_Employee.EMP_ID = cboName.ItemData(Minus2Zero(cboName.ListIndex))
      m_Employee.CURRENT_POSITION = cboPosition.ItemData(Minus2Zero(cboPosition.ListIndex))
      If m_Employee.EMP_ID <> 0 Then
           If Not glbDaily.QueryEmployeeMoney(m_Employee, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
      End If
   End If
   IsOK = True
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
        If ItemCount > 0 Then
        Call m_Employee.PopulateFromRSMoney(1, m_Rs)
        txtSalary.Text = m_Employee.CURRENT_SALARY
        txtLendReMinder.Text = m_Employee.TOTBORROW
        End If
           m_sliptSalary.EMP_ID = cboName.ItemData(Minus2Zero(cboName.ListIndex))
            m_sliptSalary.MONTH_NO = cboMonth.ItemData(Minus2Zero(cboMonth.ListIndex))
            m_sliptSalary.YEAR_NO = Val(cboYear.Text)
            If m_sliptSalary.YEAR_NO <> 0 Then
           If Not glbDaily.QuerySliptSalary(m_sliptSalary, m_Rs1, ItemCount1, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
       End If
        If ItemCount1 > 0 Then
        Call m_sliptSalary.PopulateFromRS(m_Rs1)
        txtoldSalary.Text = m_sliptSalary.SALARY
        txtSalaryAdd.Text = m_sliptSalary.SUM_SLIPT_ADD
        txtSalarySub.Text = m_sliptSalary.SUM_SLIPT_SUB
         txtTotSalary.Text = (m_sliptSalary.SALARY + m_sliptSalary.SUM_SLIPT_ADD - m_sliptSalary.SUM_SLIPT_SUB)
      txtLendReMinder.Text = m_sliptSalary.SUM_BORROW
       cmdAdd.Enabled = False
       cmdEdit.Enabled = False
       cmdDelete.Enabled = False
      GridEX1.Enabled = False
     cmdDeleteMain.Enabled = True
        Else
                
                End If
        
           Call TabStrip1_Click
   Call EnableForm(Me, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 116 Then
 '     Call cmdSearch_Click
  '    KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
      Call cmdClear_Click
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
GridEX1.Columns.Item(1).Visible = False
  Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "REAL ID"
GridEX1.Columns.Item(2).Visible = False
   Set Col = GridEX1.Columns.add '3
   Col.Width = 4900
   Col.Caption = MapText("ประเภท")
      
   Set Col = GridEX1.Columns.add '4
   Col.Width = 1730
   Col.Caption = MapText("จำนวนเงิน")
   Col.TextAlignment = jgexAlignRight

   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("สลิปเงินเดือน")
   pnlHeader.Caption = MapText("สลิปเงินเดือน")
   
   Call InitGrid
   
   Call InitNormalLabel(lblMonth, MapText("เดือน"))
   Call InitNormalLabel(lblPosition, MapText("ตำแหน่ง"))
   Call InitNormalLabel(lblYear, MapText("ปี"))
   
   Call InitNormalLabel(lblCurrentSalary, MapText("เงินเดือนปัจจุบัน"))
   Call InitNormalLabel(lbloldSalary, MapText("เงินเดือน"))
   Call InitNormalLabel(lblSalaryAdd, MapText("ส่วนเพิ่มเงินเดือน"))
   Call InitNormalLabel(lblSalarySub, MapText("ส่วนหักเงินเดือน"))
   Call InitNormalLabel(lblLendReMinder, MapText("เงินยืมรวมคงเหลือ"))
   Call InitNormalLabel(lblTotSalary, MapText("เงินสุทธิ"))
   
Call InitNormalLabel(lblBath0, MapText("บาท"))
   Call InitNormalLabel(lblBath1, MapText("บาท"))
   Call InitNormalLabel(lblBath2, MapText("บาท"))
   Call InitNormalLabel(lblBath3, MapText("บาท"))
   Call InitNormalLabel(lblBath4, MapText("บาท"))
   Call InitNormalLabel(lblBath5, MapText("บาท"))
   Call txtSalary.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.NAME_LEN)
   Call txtoldSalary.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.NAME_LEN)
   Call txtSalaryAdd.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.NAME_LEN)
   Call txtSalarySub.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.NAME_LEN)
   Call txtLendReMinder.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.NAME_LEN)
   Call txtTotSalary.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.NAME_LEN)
   txtSalary.Enabled = False
   txtSalaryAdd.Enabled = False
   txtSalarySub.Enabled = False
   txtLendReMinder.Enabled = False
   txtTotSalary.Enabled = False
   txtoldSalary.Enabled = False
   Call InitCombo(cboPosition)
   Call InitCombo(cboName)
   Call InitCombo(cboMonth)
Call InitThaiMonth(cboMonth)
    Call InitCombo(cboYear)
Call InitThaiYear(cboYear)
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDeleteMain.Picture = LoadPicture(glbParameterObj.NormalButton1)
  
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   Call InitMainButton(cmdDeleteMain, MapText("ลบสลิปเงินเดือน"))
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("ส่วนบวกเงินเดือน")
   TabStrip1.Tabs.add().Caption = MapText("ส่วนหักเงินเดือน")
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"
   
   Set m_Employee = New CEmployee
   Set m_TempEmployee = New CEmployee
   Set m_Rs = New ADODB.Recordset
   Set m_sliptSalary = New CSliptSalary
   Set m_TempSliptSalary = New CSliptSalary
   Set m_EmpReceivable = New CEmpReceivable
   Set m_Rs1 = New ADODB.Recordset

   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub



Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(4)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"
   If TabStrip1.SelectedItem.Index = 1 Then
     If m_sliptSalary.SliptAdd Is Nothing Then
         Exit Sub
      End If
      
      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim SA As CSliptAdd
           Set SA = GetItem(m_sliptSalary.SliptAdd, RowIndex, RealIndex)
      If SA Is Nothing Then
         Exit Sub
      End If
      Values(1) = SA.SLIPT_ADD_ID
      Values(2) = RealIndex
      Values(3) = SA.MONTHLY_NAME
      Values(4) = SA.MONTHLY_AMOUNT
    
    ElseIf TabStrip1.SelectedItem.Index = 2 Then
     If m_sliptSalary.SliptSub Is Nothing Then
         Exit Sub
      End If
      
      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim SB As CSliptSub
           Set SB = GetItem(m_sliptSalary.SliptSub, RowIndex, RealIndex)
      If SB Is Nothing Then
         Exit Sub
      End If
      Values(1) = SB.SLIPT_SUB_ID
      Values(2) = RealIndex
      Values(3) = SB.MONTHLY_NAME
      Values(4) = SB.MONTHLY_AMOUNT
    
End If
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip1_Click()
Call InitGrid
If TabStrip1.SelectedItem.Index = 1 Then
            If m_sliptSalary.SliptAdd Is Nothing Then
            Exit Sub
            End If
            GridEX1.ItemCount = CountItem(m_sliptSalary.SliptAdd)
      GridEX1.Rebind
         ElseIf TabStrip1.SelectedItem.Index = 2 Then
            If m_sliptSalary.SliptSub Is Nothing Then
            Exit Sub
            End If
            GridEX1.ItemCount = CountItem(m_sliptSalary.SliptSub)
      GridEX1.Rebind
End If
End Sub

Private Sub CalculateSliptAdd()
Dim II As CSliptAdd
Dim III As CSliptSub
Dim Sum As Double
Dim Remind As Double

If TabStrip1.SelectedItem.Index = 1 Then
   For Each II In m_sliptSalary.SliptAdd
      If II.Flag <> "D" Then
         Sum = Sum + II.MONTHLY_AMOUNT
      End If
   Next II
       txtSalaryAdd.Text = Sum
Else
           If TempRemind = 0 Then
           TempRemind = Val(txtLendReMinder.Text)
           End If
           txtLendReMinder.Text = TempRemind
   For Each III In m_sliptSalary.SliptSub
      If III.Flag <> "D" Then
         Sum = Sum + III.MONTHLY_AMOUNT
      If III.MONTHLY_SUB = 1 Then
      Remind = Remind + III.MONTHLY_AMOUNT
      End If
      End If
   Next III
              If Remind > Val(txtLendReMinder.Text) Then
                Call MsgBox("ไม่ควรหักเงินยืมมากกว่าเงินยืมที่เป็นจริง กรุณาแก้ไขข้อมูลใหม่", vbOKOnly, PROJECT_NAME)
                Set m_sliptSalary.SliptSub = Nothing
                Set m_sliptSalary.SliptSub = New Collection
                GridEX1.ItemCount = CountItem(m_sliptSalary.SliptSub)
        GridEX1.Rebind
        txtTotSalary.Text = Val(txtTotSalary.Text) + Val(txtSalarySub.Text)
        txtSalarySub.Text = ""
        Exit Sub
        End If
      txtLendReMinder.Text = Val(txtLendReMinder.Text) - Remind
      txtSalarySub.Text = Sum
End If
       m_sliptSalary.SALARY = m_Employee.CURRENT_SALARY
        txtoldSalary.Text = m_sliptSalary.SALARY
        m_sliptSalary.SUM_SLIPT_ADD = Val(frmSliptSalary.txtSalaryAdd.Text)
        m_sliptSalary.SUM_SLIPT_SUB = Val(frmSliptSalary.txtSalarySub.Text)
         txtTotSalary.Text = (m_sliptSalary.SALARY + m_sliptSalary.SUM_SLIPT_ADD - m_sliptSalary.SUM_SLIPT_SUB)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   m_sliptSalary.SLIPT_SALARY_ID = id
   m_sliptSalary.AddEditMode = SHOW_ADD
   m_sliptSalary.EMP_ID = cboName.ItemData(Minus2Zero(cboName.ListIndex))
   m_sliptSalary.MONTH_NO = cboMonth.ItemData(Minus2Zero(cboMonth.ListIndex))
   m_sliptSalary.YEAR_NO = Val(cboYear.Text)
  m_sliptSalary.SALARY = Val(txtoldSalary.Text)
   m_sliptSalary.SUM_SLIPT_ADD = Val(txtSalaryAdd.Text)
   m_sliptSalary.SUM_SLIPT_SUB = Val(txtSalarySub.Text)
   m_sliptSalary.SUM_BORROW = Val(txtLendReMinder.Text)
   m_sliptSalary.LOCK_FLAG = "Y"
   Call EnableForm(Me, False)
   Call glbDaily.StartTransaction
   If Not glbDaily.AddEditSliptSalary(m_sliptSalary, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   m_Employee.TOTBORROW = Val(txtLendReMinder.Text)
    If Not glbDaily.AddEditEmployeeMoney(m_Employee, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
If Val(txtLendReMinder.Text) = 0 Then
m_EmpReceivable.CLOSED_FLAG = "Y"
m_EmpReceivable.EMP_ID = cboName.ItemData(Minus2Zero(cboName.ListIndex))
m_EmpReceivable.AddEditDataFlag
End If
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Call glbDaily.RollbackTransaction
      Exit Function
   End If
   
   Call glbDaily.CommitTransaction
   Call EnableForm(Me, True)
   SaveData = True
End Function

                                
Private Sub cmdDeleteMain_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim id As Long
Dim III As CSliptSub
Dim Remind As Double
   
   If m_sliptSalary.SLIPT_SALARY_ID <= 0 Then
      Exit Sub
  End If
  
   id = m_sliptSalary.SLIPT_SALARY_ID
   
   Call glbDatabaseMngr.LockTable(m_TableName, id, IsCanLock, glbErrorLog)
   If Not ConfirmDelete("สลิปเงินเดือน") Then
      Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
      Exit Sub
   End If

   Call EnableForm(Me, False)
  Call glbDaily.StartTransaction
   If Not glbDaily.DeleteSliptSalary(id, IsOK, False, glbErrorLog) Then
     m_sliptSalary.SLIPT_SALARY_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Sub
   End If


For Each III In m_sliptSalary.SliptSub
      If III.MONTHLY_SUB = 1 Then
      Remind = Remind + III.MONTHLY_AMOUNT
      End If
      Next III
   
   m_Employee.TOTBORROW = m_Employee.TOTBORROW + Remind
 If Not glbDaily.AddEditEmployeeMoney(m_Employee, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Sub
   End If
Call glbDaily.CommitTransaction

   Call cmdClear_Click
   
   Call glbDatabaseMngr.UnLockTable(m_TableName, id, IsCanLock, glbErrorLog)
   Call EnableForm(Me, True)
End Sub

