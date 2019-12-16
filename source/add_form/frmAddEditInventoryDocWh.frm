VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddEditInventoryDocWh 
   ClientHeight    =   10965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14430
   ForeColor       =   &H00000000&
   Icon            =   "frmAddEditInventoryDocWh.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10965
   ScaleWidth      =   14430
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   16320
      ScaleHeight     =   1035
      ScaleWidth      =   555
      TabIndex        =   46
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   10920
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   19262
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboCon3 
         Height          =   315
         Left            =   3720
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ComboBox cboCon2 
         Height          =   315
         Left            =   3720
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.ComboBox cboCon1 
         Height          =   315
         Left            =   3720
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   2880
         Width           =   1455
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   5760
         Width           =   13755
         _ExtentX        =   24262
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
      Begin prjFarmManagement.uctlTextLookup uctlCustomerLookup 
         Height          =   435
         Left            =   6000
         TabIndex        =   4
         Top             =   1260
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6000
         TabIndex        =   2
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjFarmManagement.uctlTextBox txtDoNo 
         Height          =   435
         Left            =   1560
         TabIndex        =   3
         Top             =   1260
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   1560
         TabIndex        =   1
         Top             =   840
         Width           =   2295
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTruckNo 
         Height          =   435
         Left            =   1560
         TabIndex        =   5
         Top             =   1710
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3375
         Left            =   120
         TabIndex        =   18
         Top             =   6240
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   5953
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
         Column(1)       =   "frmAddEditInventoryDocWh.frx":27A2
         Column(2)       =   "frmAddEditInventoryDocWh.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditInventoryDocWh.frx":290E
         FormatStyle(2)  =   "frmAddEditInventoryDocWh.frx":2A6A
         FormatStyle(3)  =   "frmAddEditInventoryDocWh.frx":2B1A
         FormatStyle(4)  =   "frmAddEditInventoryDocWh.frx":2BCE
         FormatStyle(5)  =   "frmAddEditInventoryDocWh.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditInventoryDocWh.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   14445
         _ExtentX        =   25479
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtDesc 
         Height          =   435
         Left            =   6000
         TabIndex        =   6
         Top             =   1710
         Width           =   5385
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtEntryWeight 
         Height          =   435
         Left            =   8040
         TabIndex        =   33
         Top             =   2880
         Width           =   1995
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtExitWeight 
         Height          =   435
         Left            =   8040
         TabIndex        =   34
         Top             =   3360
         Width           =   1995
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeightAmount 
         Height          =   435
         Left            =   8040
         TabIndex        =   35
         Top             =   3840
         Width           =   1995
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeightNote 
         Height          =   435
         Left            =   8040
         TabIndex        =   41
         Top             =   4320
         Width           =   4395
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlEmpCheckCar 
         Height          =   435
         Left            =   3720
         TabIndex        =   10
         Top             =   4800
         Width           =   5025
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextLookup uctlEmpCheckProduct 
         Height          =   435
         Left            =   3720
         TabIndex        =   11
         Top             =   5280
         Width           =   5025
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtWeightNo 
         Height          =   435
         Left            =   8040
         TabIndex        =   49
         Top             =   2400
         Width           =   1995
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtEntryWeightTime 
         Height          =   435
         Left            =   12120
         TabIndex        =   52
         Top             =   2880
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtExitWeightTime 
         Height          =   435
         Left            =   12120
         TabIndex        =   53
         Top             =   3360
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtExitWeightDate 
         Height          =   435
         Left            =   10800
         TabIndex        =   54
         Top             =   3360
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtEntryWeightDate 
         Height          =   435
         Left            =   10800
         TabIndex        =   55
         Top             =   2880
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
      End
      Begin Threed.SSCheck sscConsignment 
         Height          =   495
         Left            =   9720
         TabIndex        =   56
         Top             =   5280
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   131073
         Caption         =   "sscConsignment"
      End
      Begin VB.Label lblWeightTime 
         Alignment       =   2  'Center
         Caption         =   "---"
         Height          =   315
         Left            =   10800
         TabIndex        =   51
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label lblWeightNo 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   315
         Left            =   5760
         TabIndex        =   50
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label lblCheckProduct 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   315
         Left            =   0
         TabIndex        =   48
         Top             =   5400
         Width           =   3495
      End
      Begin VB.Label lblCheckCar 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   315
         Left            =   0
         TabIndex        =   47
         Top             =   4920
         Width           =   3495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   315
         Left            =   10080
         TabIndex        =   45
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   315
         Left            =   10080
         TabIndex        =   44
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   315
         Left            =   10080
         TabIndex        =   43
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label lblNote 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   315
         Left            =   5760
         TabIndex        =   42
         Top             =   4440
         Width           =   2055
      End
      Begin Threed.SSCommand cmdLoadWeight 
         Height          =   405
         Left            =   11040
         TabIndex        =   40
         Top             =   3840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDocWh.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblCon3 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   315
         Left            =   0
         TabIndex        =   39
         Top             =   3960
         Width           =   3495
      End
      Begin VB.Label lblCon2 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   315
         Left            =   120
         TabIndex        =   38
         Top             =   3480
         Width           =   3375
      End
      Begin VB.Label lblCon1 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   315
         Left            =   240
         TabIndex        =   37
         Top             =   3000
         Width           =   3255
      End
      Begin VB.Label lblConLog 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   315
         Left            =   240
         TabIndex        =   36
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblWeightTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   315
         Left            =   5760
         TabIndex        =   32
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label lblWeightOut 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   315
         Left            =   5760
         TabIndex        =   31
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label lblWeightIn 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   315
         Left            =   5760
         TabIndex        =   30
         Top             =   3000
         Width           =   2055
      End
      Begin Threed.SSCommand cmdOther 
         Height          =   525
         Left            =   5100
         TabIndex        =   19
         Top             =   9840
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDocWh.frx":3250
         ButtonStyle     =   3
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2730
         TabIndex        =   28
         Top             =   1800
         Width           =   435
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   6720
         TabIndex        =   15
         Top             =   9840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDocWh.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   3870
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDocWh.frx":3884
         ButtonStyle     =   3
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4470
         TabIndex        =   27
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Label lblCustomerNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4470
         TabIndex        =   26
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4710
         TabIndex        =   25
         Top             =   870
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8400
         TabIndex        =   16
         Top             =   9840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDocWh.frx":3B9E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10080
         TabIndex        =   17
         Top             =   9840
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1800
         TabIndex        =   13
         Top             =   9840
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   120
         TabIndex        =   12
         Top             =   9840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDocWh.frx":3EB8
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3480
         TabIndex        =   14
         Top             =   9840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditInventoryDocWh.frx":41D2
         ButtonStyle     =   3
      End
      Begin VB.Label lblTruckNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -120
         TabIndex        =   23
         Top             =   1740
         Width           =   1575
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -150
         TabIndex        =   22
         Top             =   900
         Width           =   1665
      End
      Begin VB.Label lblDoNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -60
         TabIndex        =   21
         Top             =   1320
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmAddEditInventoryDocWh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_InventoryWHDoc As CInventoryWHDoc
Private m_LotItemWh As CLotItemWH
Private m_Weight As CWeight
Private m_CollLotDoc As Collection
Private m_CollJob As Collection
Private m_Customers As Collection
Private m_Locations As Collection
Private m_Employees As Collection
Private m_Employees2 As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public CancelWeigth As Boolean
Public AutoGenPo As Boolean

Public ID As Long
Public ID2 As Long
Public DocumentType As Long

Private FileName As String
Private m_Cd As Collection
Private TempWeight As Collection
Private CW As CWeight
Private DocAdd As Long
Public CustomerID As Long
Public DOCUMENT_NO As String
Public DOCUMENT_ID_RQ  As Long
Public TRUCK_NO As String
Private NotCheckWeigth As Boolean
Private EditWeigth As Boolean
Private TempWeigth As String
Private iniNotCheckWeigth As String
Private LocationIm As Long

Private m_InventoryDoc As CInventoryDoc
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
   
   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_InventoryWHDoc.INVENTORY_WH_DOC_ID = ID
      m_InventoryWHDoc.COMMIT_FLAG = ""
      m_InventoryWHDoc.QueryFlag = 1
      
      If Not glbDaily.QueryInventoryWhDocForLG(m_InventoryWHDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
        Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_InventoryWHDoc.PopulateFromRS(1, m_Rs)
      uctlDocumentDate.ShowDate = m_InventoryWHDoc.DOCUMENT_DATE
      txtDoNo.Text = m_InventoryWHDoc.DO_NO
      txtTruckNo.Text = m_InventoryWHDoc.TRUCK_NO
      txtDocumentNo.Text = m_InventoryWHDoc.DOCUMENT_NO
      uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_InventoryWHDoc.CUSTOMER_ID)
      txtDesc.Text = m_InventoryWHDoc.NOTE
      cboCon1.ListIndex = m_InventoryWHDoc.CONDITION1
      cboCon2.ListIndex = m_InventoryWHDoc.CONDITION2
      cboCon3.ListIndex = m_InventoryWHDoc.CONDITION3
      uctlEmpCheckCar.MyCombo.ListIndex = IDToListIndex(uctlEmpCheckCar.MyCombo, m_InventoryWHDoc.EMP_CHECK_CAR_ID)
      uctlEmpCheckProduct.MyCombo.ListIndex = IDToListIndex(uctlEmpCheckProduct.MyCombo, m_InventoryWHDoc.EMP_CHECK_PRODUCT_ID)
      txtEntryWeight.Text = IIf(m_InventoryWHDoc.ENTRY_WEIGHT = 0, "", m_InventoryWHDoc.ENTRY_WEIGHT)
      txtExitWeight.Text = IIf(m_InventoryWHDoc.EXIT_WEIGHT = 0, "", m_InventoryWHDoc.EXIT_WEIGHT)
      txtWeightAmount.Text = IIf(m_InventoryWHDoc.EXIT_WEIGHT = 0, "", m_InventoryWHDoc.TOTAL_WEIGHT)
      txtWeightNote.Text = m_InventoryWHDoc.WEIGHT_NOTE
      txtWeightNo.Text = m_InventoryWHDoc.WEIGHT_ID
      txtEntryWeightDate.Text = m_InventoryWHDoc.ENTRY_WEIGHT_DATE
      txtEntryWeightTime.Text = m_InventoryWHDoc.ENTRY_WEIGHT_TIME 'TimeToStringHHMM(m_InventoryWHDoc.ENTRY_WEIGHT_TIME)
      txtExitWeightDate.Text = m_InventoryWHDoc.EXIT_WEIGHT_DATE
      txtExitWeightTime.Text = m_InventoryWHDoc.EXIT_WEIGHT_TIME
      sscConsignment.Value = FlagToCheck(m_InventoryWHDoc.CONSIGNMENT_FLAG)
      
      DOCUMENT_ID_RQ = m_InventoryWHDoc.INVENTORY_DOC_ID
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



Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim IvdWH As CInventoryWHDoc
Dim Sp As CSupItem
Dim LtWh As CLotItemWH
Dim StrStockAmount As String
Dim TempDocNo As String
Dim firstDate As Date
Dim lastDate As Date
Dim MonthlyAccums  As Collection
Dim YYYYMM As String
Dim BalanceLi As CLotItem
Dim TempLi1 As CLotItem
Dim TempLi2 As CLotItem
Dim InventoryBals1  As Collection
   
   If Not VerifyTextControl(lblDocumentNo, txtDocumentNo, False) Then
      Exit Function
   End If
   
   If Not VerifyDate(lblDocumentDate, uctlDocumentDate, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblCustomerNo, uctlCustomerLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblTruckNo, txtTruckNo, False) Then
      Exit Function
   End If
   
   If Not NotCheckWeigth Then
      If Not VerifyTextControl(lblWeightIn, txtEntryWeight, False) Then
         Exit Function
      End If
   End If
   
   If m_InventoryWHDoc.INVENTORY_DOC_ID > 0 Then
      glbErrorLog.LocalErrorMsg = MapText("มีการออกใบฝากขายแล้ว ไม่สามารถแก้ไขเอกสารนี้ได้")
      glbErrorLog.ShowUserError
      Exit Function
   End If

   If ShowMode = SHOW_ADD Then
      If Not CheckUniqueNs(INVENTORY_WH_DOC_UNIQUE, txtDocumentNo.Text, -1) Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         txtDocumentNo.Text = ""
         Call LoadConfigDoc(Nothing, m_Cd)
         DocAdd = DocAdd + 1
         Call cmdAuto_Click
         Exit Function
      End If
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   

   If Not TempWeigth = iniNotCheckWeigth Or Not EditWeigth Then  ' ถ้า เช็คว่า น้ำหนักเดิมเป็นการใช้น้ำหนักชั่วคราว และได้เข้ามา update น้ำหนักใหม่  หรือเป็นการแก้ไขน้ำหนักใหม่กรณีออกใบส่งของแล้ว ก็ให้ผ่านการตรวจสอบไป เพื่อให้สามารถแก้ไขได้
      If glbUser.USER_NAME <> "ADMIN" Then
         If ShowMode = SHOW_EDIT Then
            If m_InventoryWHDoc.LOAD_FLAG = "Y" Then
               glbErrorLog.LocalErrorMsg = MapText("เอกสาร ") & " " & txtDocumentNo.Text & " " & MapText("ออกใบส่งของแล้วไม่สามารถแก้ไขได้ หากต้องการแก้ไขให้ไปแก้ไขใบส่งของก่อน")
               glbErrorLog.ShowUserError
               SaveData = True
               Exit Function
            End If
         End If
      End If
      
   m_InventoryWHDoc.AddEditMode = ShowMode
   m_InventoryWHDoc.INVENTORY_WH_DOC_ID = ID
   m_InventoryWHDoc.DOCUMENT_DATE = uctlDocumentDate.ShowDate
   m_InventoryWHDoc.TRUCK_NO = txtTruckNo.Text
   m_InventoryWHDoc.DOCUMENT_NO = txtDocumentNo.Text
   m_InventoryWHDoc.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   m_InventoryWHDoc.DOCUMENT_TYPE = DocumentType
   m_InventoryWHDoc.NOTE = txtDesc.Text
   m_InventoryWHDoc.ENTRY_WEIGHT = Val(txtEntryWeight.Text) 'FormatNumber(Val(txtEntryWeight.Text))
   m_InventoryWHDoc.EXIT_WEIGHT = Val(txtExitWeight.Text)
   m_InventoryWHDoc.TOTAL_WEIGHT = Val(txtWeightAmount.Text)
   m_InventoryWHDoc.WEIGHT_NOTE = txtWeightNote.Text
   
   m_InventoryWHDoc.WEIGHT_ID = txtWeightNo.Text
   m_InventoryWHDoc.ENTRY_WEIGHT_DATE = txtEntryWeightDate.Text
   m_InventoryWHDoc.ENTRY_WEIGHT_TIME = txtEntryWeightTime.Text
   m_InventoryWHDoc.EXIT_WEIGHT_DATE = txtExitWeightDate.Text
   m_InventoryWHDoc.EXIT_WEIGHT_TIME = txtExitWeightTime.Text
   
   m_InventoryWHDoc.NOTE = txtDesc.Text
   m_InventoryWHDoc.CONDITION1 = cboCon1.ItemData(Minus2Zero(cboCon1.ListIndex))
   m_InventoryWHDoc.CONDITION2 = cboCon2.ItemData(Minus2Zero(cboCon2.ListIndex))
   m_InventoryWHDoc.CONDITION3 = cboCon3.ItemData(Minus2Zero(cboCon3.ListIndex))
   m_InventoryWHDoc.EMP_CHECK_CAR_ID = uctlEmpCheckCar.MyCombo.ItemData(Minus2Zero(uctlEmpCheckCar.MyCombo.ListIndex))
   m_InventoryWHDoc.EMP_CHECK_PRODUCT_ID = uctlEmpCheckProduct.MyCombo.ItemData(Minus2Zero(uctlEmpCheckProduct.MyCombo.ListIndex))
   m_InventoryWHDoc.EXCEPTION_FLAG = "N"
   m_InventoryWHDoc.SUCCESS_FLAG = "N"
   m_InventoryWHDoc.DO_NO = txtDoNo.Text
   m_InventoryWHDoc.CONSIGNMENT_FLAG = Check2Flag(sscConsignment.Value)

   For Each m_LotItemWh In m_InventoryWHDoc.C_LotItemsWH 'ตรวจสอบกรณี ลบรายการ
      If m_LotItemWh.Flag = "D" Then
         m_InventoryWHDoc.LOAD_FLAG = "L" 'รอเพิ่มรายการอาหาร
      End If
   Next m_LotItemWh

   For Each m_LotItemWh In m_InventoryWHDoc.C_LotItemsWH
      If m_LotItemWh.LOAD_AMOUNT_FLAG = "N" Or m_LotItemWh.LOAD_AMOUNT_FLAG = "" Then 'ถ้ายังมีสินค้าบางตัวยัง โหลดน้ำหนักไม่เรียบร้อย
         m_InventoryWHDoc.LOAD_FLAG = "N" 'รอขึ้นอาหาร
         Exit For
      Else
         m_InventoryWHDoc.LOAD_FLAG = "C" 'รอชั่งน้ำหนักออก
      End If
   Next m_LotItemWh

   If (Val(m_InventoryWHDoc.EXIT_WEIGHT) = 0 Or Val(m_InventoryWHDoc.EXIT_WEIGHT) = iniNotCheckWeigth) And (m_InventoryWHDoc.LOAD_FLAG = "C") Then
      m_InventoryWHDoc.LOAD_FLAG = "C" 'รอชั่งน้ำหนักออก
      If Not DocumentType = 2004 Then
         If Not VerifyCombo(lblCon1, cboCon1, False) Then
            Exit Function
         End If
         If Not VerifyCombo(lblCon2, cboCon2, False) Then
            Exit Function
         End If
         If Not VerifyCombo(lblCon3, cboCon3, False) Then
            Exit Function
         End If
         If Not VerifyCombo(lblCheckCar, uctlEmpCheckCar.MyCombo, False) Then
            Exit Function
         End If
         If Not VerifyCombo(lblCheckProduct, uctlEmpCheckProduct.MyCombo, False) Then
            Exit Function
         End If
      End If
   End If

   If Not NotCheckWeigth Then 'กรณีเปิดบิลก่อนรถเข้ามาจริง โปรแกรมจะยอมให้ปล่อยน้ำหนัก เข้าออก ได้
      If (Val(m_InventoryWHDoc.EXIT_WEIGHT) > 0) And (m_InventoryWHDoc.LOAD_FLAG = "C") Then
         m_InventoryWHDoc.LOAD_FLAG = "I" 'รอออกใบส่งของ
      End If
   Else
     If (m_InventoryWHDoc.LOAD_FLAG = "C") Then
         m_InventoryWHDoc.LOAD_FLAG = "I" 'รอออกใบส่งของ
     End If
   End If

   
   If m_InventoryWHDoc.LOAD_FLAG = "" Then 'ถ้ายังไม่มีรายการ
      m_InventoryWHDoc.LOAD_FLAG = "A" 'รอเพิ่มรายการจาก SO
   End If
   
   If CountItem(m_InventoryWHDoc.C_LotItemsWH) = 0 Then
      m_InventoryWHDoc.LOAD_FLAG = "A" 'รอเพิ่มรายการจาก SO
   End If
      
   Call EnableForm(Me, False)
   Call glbDaily.StartTransaction
   
   If Not glbDaily.AddEditInventoryWhDoc(m_InventoryWHDoc, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   

   If m_InventoryWHDoc.INVENTORY_DOC_ID <= 0 And m_InventoryWHDoc.LOAD_FLAG = "I" Then
      If sscConsignment.Value = ssCBChecked Then
           '  สร้าง เอกสารการโอน
         Call CreateDocTranfer
      End If
   End If
   
   'Update สถานะเอกสาร ของ BillingDoc หาก ดึงใบ so แล้วก็ให้เปลี่ยนเป็น Y
   Dim BD As CBillingDoc
   Dim PrevKey As Long
   Set BD = New CBillingDoc
   For Each m_LotItemWh In m_InventoryWHDoc.C_LotItemsWH
     If PrevKey <> m_LotItemWh.BILLING_DOC_ID Then
         BD.BILLING_DOC_ID = m_LotItemWh.BILLING_DOC_ID
         PrevKey = m_LotItemWh.BILLING_DOC_ID
         If m_InventoryWHDoc.LOAD_FLAG = "I" Then
            Call BD.UpdateSuccessFlag("Y")
         Else
            Call BD.UpdateSuccessFlag("N")
         End If
      End If
   Next m_LotItemWh
   Set BD = Nothing
   
   'ตรวจสอบ Stock หลังจาก update ใหม่ เพื่อเปลี่ยนสถานะ ของ Out Stock Flag
   For Each m_LotItemWh In m_InventoryWHDoc.C_LotItemsWH
     If m_LotItemWh.Flag = "D" Then
         Call LoadLotInPartIemAmount(Nothing, Nothing, , , , , m_LotItemWh.PART_ITEM_ID, 2, 1, 1, "I", m_LotItemWh.C_LotDoc, , DocumentType, m_LotItemWh.Flag)
    End If
   Next m_LotItemWh
   
Else 'ให้แก้ไขเฉพาะน้ำหนักเท่านั้น

   Call EnableForm(Me, False)
   Call glbDaily.StartTransaction
   
      m_InventoryWHDoc.AddEditMode = ShowMode
      m_InventoryWHDoc.INVENTORY_WH_DOC_ID = ID
      m_InventoryWHDoc.ENTRY_WEIGHT = Val(txtEntryWeight.Text)
      m_InventoryWHDoc.EXIT_WEIGHT = Val(txtExitWeight.Text)
      m_InventoryWHDoc.TOTAL_WEIGHT = Val(txtWeightAmount.Text)
      m_InventoryWHDoc.WEIGHT_NOTE = txtWeightNote.Text
      m_InventoryWHDoc.WEIGHT_ID = txtWeightNo.Text
      m_InventoryWHDoc.ENTRY_WEIGHT_DATE = txtEntryWeightDate.Text
      m_InventoryWHDoc.ENTRY_WEIGHT_TIME = txtEntryWeightTime.Text
      m_InventoryWHDoc.EXIT_WEIGHT_DATE = txtExitWeightDate.Text
      m_InventoryWHDoc.EXIT_WEIGHT_TIME = txtExitWeightTime.Text
      
      IsOK = m_InventoryWHDoc.UpdateWeight
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
Function CreateDocTranfer()
Dim DocumentNo As String
Dim TempCollection As Collection

Set m_InventoryDoc = New CInventoryDoc

' LocationIm = uctlPlaceLookup.MyCombo.ItemData(Minus2Zero(uctlPlaceLookup.MyCombo.ListIndex))
   m_InventoryDoc.AddEditMode = SHOW_ADD
   m_InventoryDoc.INVENTORY_DOC_ID = -1
    m_InventoryDoc.DOCUMENT_DATE = uctlDocumentDate.ShowDate

   Call glbDatabaseMngr.GenerateNumber(TRANSFER_RQ_NUMBER, DocumentNo, glbErrorLog)

   m_InventoryDoc.DOCUMENT_NO = DocumentNo
   m_InventoryDoc.DELIVERY_FEE = 0
   m_InventoryDoc.EMP_ID = -1 'uctlEmployeeLookup.MyCombo.ItemData(Minus2Zero(uctlEmployeeLookup.MyCombo.ListIndex))
   m_InventoryDoc.DOCUMENT_TYPE = 3
   m_InventoryDoc.EXCEPTION_FLAG = "Y"
   m_InventoryDoc.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
'   m_InventoryDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   m_InventoryDoc.INVENTORY_WH_DOC_ID = m_InventoryWHDoc.INVENTORY_WH_DOC_ID
   
   Dim LIW As CLotItemWH
   If Not CountItem(m_InventoryWHDoc.C_LotItemsWH) > 0 Then
'      Set LIW = m_InventoryWHDoc.C_LotItemsWH.Item(1)
'   Else
     Exit Function
   End If
   
   
   Dim EnpAddress As CTransferItem
   Dim Ei As CLotItem
   Dim II As CLotItem
   
   For Each LIW In m_InventoryWHDoc.C_LotItemsWH
      Set Ei = New CLotItem
      Set II = New CLotItem
      Set TempCollection = New Collection
      Set EnpAddress = New CTransferItem

      Ei.Flag = "A"
      Ei.CALCULATE_FLAG = "Y"
      II.Flag = "A"
      II.CALCULATE_FLAG = "Y"
      EnpAddress.Flag = "A"

      Set EnpAddress.ExportItem = Ei
      Set EnpAddress.ImportItem = II

      Call TempCollection.add(EnpAddress)

''จ่ายออก จาก โกดังอาหาร
   EnpAddress.ExportItem.PART_TYPE = LIW.PART_TYPE 'uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   EnpAddress.ExportItem.PART_ITEM_ID = LIW.PART_ITEM_ID ' uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   EnpAddress.ExportItem.LOCATION_ID = LIW.LOCATION_ID ' uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   EnpAddress.ExportItem.TX_AMOUNT = LIW.TX_AMOUNT '  txtQuantity.Text
   EnpAddress.ExportItem.INCLUDE_UNIT_PRICE = LIW.INCLUDE_UNIT_PRICE  ' Val(txtPrice.Text)
   EnpAddress.ExportItem.PART_TYPE_NAME = LIW.PART_TYPE_NAME ' uctlPartLookup.MyCombo.Text
   EnpAddress.ExportItem.LOCATION_NAME = LIW.LOCATION_NAME  'uctlLocationLookup.MyCombo.Text
   EnpAddress.ExportItem.PART_NO = LIW.PART_NO  'uctlPartLookup.MyTextBox.Text
   EnpAddress.ExportItem.PART_DESC = LIW.PART_DESC ' uctlPartLookup.MyCombo.Text
   EnpAddress.ExportItem.TX_TYPE = "E"
   EnpAddress.ExportItem.PACKAGING_AMT = LIW.PACKAGING_AMT  'Val(txtPackaging.Text)

''******************
'
''รับเข้าโกดัง หมอน้อง
   EnpAddress.ImportItem.PART_TYPE = LIW.PART_TYPE 'uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   EnpAddress.ImportItem.PART_ITEM_ID = LIW.PART_ITEM_ID 'uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   EnpAddress.ImportItem.LOCATION_ID = LocationIm
   EnpAddress.ImportItem.LOCATION_NAME = LocationIm 'uctlPlaceLookup.MyCombo.Text
   EnpAddress.ImportItem.TX_AMOUNT = LIW.TX_AMOUNT ' txtQuantity.Text
   EnpAddress.ImportItem.ACTUAL_UNIT_PRICE = 0 'Val(txtPrice.Text)
   EnpAddress.ImportItem.TOTAL_ACTUAL_PRICE = LIW.TOTAL_ACTUAL_PRICE  ' (txtQuantity.Text) * Val(txtPrice.Text)
   EnpAddress.ImportItem.INCLUDE_UNIT_PRICE = LIW.INCLUDE_UNIT_PRICE  'Val(txtPrice.Text)
   EnpAddress.ImportItem.TOTAL_INCLUDE_PRICE = LIW.TOTAL_INCLUDE_PRICE ' EnpAddress.ImportItem.TOTAL_ACTUAL_PRICE
   EnpAddress.ImportItem.PART_TYPE_NAME = LIW.PART_TYPE_NAME '  uctlPartLookup.MyCombo.Text
   EnpAddress.ImportItem.PART_NO = LIW.PART_NO   'uctlPartLookup.MyTextBox.Text
   EnpAddress.ImportItem.PART_DESC = LIW.PART_DESC 'uctlPartLookup.MyCombo.Text
   EnpAddress.ImportItem.TX_TYPE = "I"
   EnpAddress.ImportItem.LAYOUT_ID = LIW.LAYOUT_ID  ' uctlLayoutLookup.MyCombo.ItemData(Minus2Zero(uctlLayoutLookup.MyCombo.ListIndex))
   EnpAddress.ImportItem.PACKAGING_AMT = LIW.PACKAGING_AMT ' Val(txtPackaging.Text)
'   '******************
   
   Call m_InventoryDoc.TransferItems.add(EnpAddress)

Next LIW

   Call CreateImportExportItems
   If (m_InventoryDoc.COMMIT_FLAG = "Y") Then
      If m_InventoryDoc.OLD_COMMIT_FLAG <> "Y" Then
         Call glbDaily.TriggerCommit(m_InventoryDoc.ImportExports)
         If Not glbDaily.VerifyStockBalance(m_InventoryDoc.ImportExports, glbErrorLog) Then
            Call EnableForm(Me, True)
            Exit Function
         End If
      End If
   End If

   If Not glbDaily.AddEditInventoryDoc(m_InventoryDoc, True, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Function
   End If
   
   
   
'   Dim IWD As CInventoryWHDoc
'   Dim tempIWD As Collection
'   Dim LTD As CLotDoc
'   Dim PD As CPalletDoc
'
'   Dim Tiwd As CInventoryWHDoc
'   Dim tLIW As CLotItemWH
'   Dim Tltd As CLotDoc
'   Dim tPD As CPalletDoc
'
'      Set Tiwd = New CInventoryWHDoc
'      Tiwd.AddEditMode = SHOW_ADD
'      Tiwd.Flag = "A"
'      Tiwd.DOCUMENT_DATE = uctlDocumentDate.ShowDate
'      Tiwd.DOCUMENT_NO = DocumentNo
'      Tiwd.START_DATE = m_InventoryWHDoc.START_DATE
'      Tiwd.FROM_DATE = m_InventoryWHDoc.FROM_DATE
'      Tiwd.FINISH_DATE = m_InventoryWHDoc.FINISH_DATE
'      If DocumentType = 2000 Then
'         Tiwd.DOCUMENT_TYPE = 20
'      ElseIf DocumentType = 2001 Then
'         Tiwd.DOCUMENT_TYPE = 21
'      End If
'
'      For Each LIW In m_InventoryWHDoc.C_LotItemsWH
'         Set tLIW = New CLotItemWH
'          tLIW.AddEditMode = SHOW_ADD
'          tLIW.Flag = "A"
'          tLIW.LOT_ITEM_WH_ID = LIW.LOT_ITEM_WH_ID
'          tLIW.BIN_NO = LIW.BIN_NO
'          tLIW.HEAD_PACK_NO = LIW.HEAD_PACK_NO
'          tLIW.LOT_ID = LIW.LOT_ID
'          tLIW.LOCK_NO = LIW.LOCK_NO
'          tLIW.PACK_AMOUNT = LIW.PACK_AMOUNT
'          tLIW.PART_ITEM_ID = LIW.PART_ITEM_ID
'          tLIW.PART_NO = LIW.PART_NO
'          tLIW.PART_DESC = LIW.PART_DESC
'          tLIW.LOCATION_ID = LocationIm
'          tLIW.PRODUCT_TYPE_ID = LIW.PRODUCT_TYPE_ID
'          tLIW.START_DATE = uctlDocumentDate.ShowDate 'LIW.START_DATE
'          tLIW.PACK_DATE = uctlDocumentDate.ShowDate
'          tLIW.TIME_PACK_BEGIN = Format(HOUR(Now), "00") & ":" & Format(Minute(Now), "00")
'          tLIW.TIME_PACK_END = Format(HOUR(Now), "00") & ":" & Format(Minute(Now), "00")
'          tLIW.TX_AMOUNT = LIW.TX_AMOUNT
'          tLIW.WEIGHT_AMOUNT = LIW.WEIGHT_AMOUNT
'          tLIW.WEIGHT_PER_PACK = LIW.WEIGHT_PER_PACK
'          tLIW.TX_TYPE = "I"
'
'            For Each LTD In LIW.C_LotDoc
'              Set Tltd = New CLotDoc
'              Tltd.AddEditMode = SHOW_ADD
'              Tltd.Flag = "A"
'              Tltd.LOT_ID = LTD.LOT_ID
'              Tltd.LOCK_NO = LTD.LOCK_NO
'
'               For Each PD In LTD.C_PalletDoc
'                  Set tPD = New CPalletDoc
'                  tPD.AddEditMode = SHOW_ADD
'                  tPD.Flag = "A"
'                  tPD.TX_TYPE = "I"
'                  tPD.CAPACITY_AMOUNT = PD.CAPACITY_AMOUNT
'                  tPD.PALLET_DOC_NO = PD.PALLET_DOC_NO
'                  Call Tltd.C_PalletDoc.add(tPD)
'                  Set tPD = Nothing
'               Next PD
'               Call tLIW.C_LotDoc.add(Tltd)
'               Set Tltd = Nothing
'            Next LTD
'         Call Tiwd.C_LotItemsWH.add(tLIW)
'         Set tLIW = Nothing
'      Next LIW
'
'
'   If Not glbDaily.AddEditInventoryWhDoc(Tiwd, False, False, glbErrorLog) Then
'      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
''      SaveData = False
'      Call glbDaily.RollbackTransaction
'      Call EnableForm(Me, True)
'      Exit Function
'   End If
'
'   Set Tiwd = Nothing
End Function
Private Sub CreateImportExportItems()
Dim Ti As CTransferItem
Dim Ei As CLotItem
Dim II As CLotItem

   Set m_InventoryDoc.ImportExports = Nothing
   Set m_InventoryDoc.ImportExports = New Collection

   For Each Ti In m_InventoryDoc.TransferItems
      Set Ei = Ti.ExportItem
      Set II = Ti.ImportItem

      Ei.Flag = Ti.Flag
      II.Flag = Ti.Flag

      Call m_InventoryDoc.ImportExports.add(Ei)
      Call m_InventoryDoc.ImportExports.add(II)
   Next Ti
End Sub
Private Sub cboCon1_Click()
m_HasModify = True
End Sub
Private Sub cboCon2_Click()
m_HasModify = True
End Sub
Private Sub cboCon3_Click()
m_HasModify = True
End Sub
Private Sub cboCondition_Click()
m_HasModify = True
End Sub

Private Sub cboDepartment_Click()
   m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim oMenu  As cPopupMenu
Dim lMenuChosen As Long
Dim LIW As CLotItemWH
Dim LTD As CLotDoc

OKClick = False
   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyCombo(lblCustomerNo, uctlCustomerLookup.MyCombo) Then
      Exit Sub
   End If
   
   If TabStrip1.SelectedItem.Index = 1 Then
         Set oMenu = New cPopupMenu
         If AutoGenPo Then
            lMenuChosen = oMenu.AddMenu(glbGuiConfigs.LoadGoodsAddMenuItems)
         Else
            lMenuChosen = oMenu.AddMenu(glbGuiConfigs.LoadGoodsAddMenuItems)
         End If
         Set oMenu = Nothing
         If lMenuChosen = 0 Then
            Exit Sub
         End If
   End If
If lMenuChosen = 1 Then
         Set frmAddSOItem.TempCollection = m_InventoryWHDoc.C_LotItemsWH
         frmAddSOItem.Area = 1 'ใบ Sale Order
                  
         Dim Di As CDoItem
         Dim TempCol As Collection
         Set TempCol = New Collection
         For Each LIW In m_InventoryWHDoc.C_LotItemsWH
             If LIW.Flag = "A" Then
                  glbErrorLog.LocalErrorMsg = MapText("กรุณาบันทึกข้อมูลก่อน")
                 glbErrorLog.ShowUserError
                 Exit Sub
             End If
            Set Di = New CDoItem
            Call Di.CopyObjectFromLtWH(1, LIW)
            If DocumentType = 2004 Then
               Call TempCol.add(Di, Trim(str(Di.BILLING_DOC_ID) & "-" & str(Di.PART_ITEM_ID) & "-" & str(Di.FEATURE_ID)))
            Else
               Call TempCol.add(Di, Trim(str(Di.BILLING_DOC_ID) & "-" & str(Di.PART_ITEM_ID)))
            End If
            Set Di = Nothing
         Next LIW
                  
         Set frmAddSOItem.m_TempCol2 = TempCol
         frmAddSOItem.DocumentDate = uctlDocumentDate.ShowDate
         frmAddSOItem.CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
         frmAddSOItem.CustomerCode = uctlCustomerLookup.MyTextBox.Text
         frmAddSOItem.DOCUMENT_TYPE = DocumentType
         frmAddSOItem.ShowMode = ShowMode 'SHOW_ADD
         If DocumentType = 2000 Then
            frmAddSOItem.HeaderText = MapText("เพิ่มรายการใบขึ้นอาหารจากใบ SALE ORDER สินค้า BAG")
         ElseIf DocumentType = 2001 Then
            frmAddSOItem.HeaderText = MapText("เพิ่มรายการใบขึ้นอาหารจากใบ SALE ORDER สินค้า BULK")
         ElseIf DocumentType = 2004 Then
            frmAddSOItem.HeaderText = MapText("เพิ่มรายการใบขึ้นอาหารจากใบ SALE ORDER ขาย อื่นๆ")
         End If
         
         Load frmAddSOItem
         frmAddSOItem.Show 1
   
         OKClick = frmAddSOItem.OKClick
         txtDesc.Text = frmAddSOItem.NOTE
   
         Unload frmAddSOItem
         Set frmAddSOItem = Nothing
         

         If DocumentType = 2000 Or DocumentType = 2001 Then
            For Each LIW In m_InventoryWHDoc.C_LotItemsWH
               If LIW.Flag <> "D" Then
                  Call LoadLotFIFOByPartItem(Nothing, m_CollLotDoc, , -1, uctlDocumentDate.ShowDate, , LIW.PART_ITEM_ID, 2, 1, 1, "I", LIW.C_LotDoc, LIW.PACK_AMOUNT, DocumentType, LIW)
                  For Each LTD In m_CollLotDoc
                     LTD.AddEditMode = SHOW_ADD
                     LTD.Flag = "A"
                     LIW.INVENTORY_WH_DOC_ID = m_InventoryWHDoc.INVENTORY_WH_DOC_ID
                     Call LIW.C_LotDoc.add(LTD)
                  Next LTD
               End If
            Next LIW
         ElseIf DocumentType = 2004 Then
              For Each LIW In m_InventoryWHDoc.C_LotItemsWH
               If LIW.Flag <> "D" Then
                   LIW.LOAD_TRUE = LIW.TX_AMOUNT
                   LIW.LOAD_AMOUNT_FLAG = "Y"
               End If
            Next LIW
         End If
         
         If OKClick Then
            GridEX1.ItemCount = CountItem(m_InventoryWHDoc.C_LotItemsWH)
            GridEX1.Rebind
         End If
ElseIf lMenuChosen = 3 Then
         m_InventoryWHDoc.AddEditMode = ShowMode
'         ShowMode = SHOW_ADD
         Set frmAddSOItem.TempCollection = m_InventoryWHDoc.C_LotItemsWH
         frmAddSOItem.Area = 3 'ใบส่งของ
         frmAddSOItem.DocumentNo = Trim(txtDoNo.Text)
         frmAddSOItem.DocumentDate = uctlDocumentDate.ShowDate
         frmAddSOItem.CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
         frmAddSOItem.CustomerCode = uctlCustomerLookup.MyTextBox.Text
         frmAddSOItem.TruckNo = Trim(txtTruckNo.Text)
         frmAddSOItem.ShowMode = ShowMode 'SHOW_ADD
         frmAddSOItem.HeaderText = MapText("เพิ่มรายการใบขึ้นอาหารจากใบ INVOICE")
         
         Load frmAddSOItem
         frmAddSOItem.Show 1
   
         OKClick = frmAddSOItem.OKClick
   
         Unload frmAddSOItem
         Set frmAddSOItem = Nothing
         
         If OKClick Then
            
            GridEX1.ItemCount = CountItem(m_InventoryWHDoc.C_LotItemsWH)
            GridEX1.Rebind
         End If
   ElseIf lMenuChosen = 5 Then
            If Not VerifyGrid(GridEX1.Value(1)) Then
               Exit Sub
            End If
            ID2 = GridEX1.Value(2)
            frmAddEditLoadGoods.ID = ID2
            Set frmAddEditLoadGoods.TempLotItemsWH = m_InventoryWHDoc.C_LotItemsWH
            frmAddEditLoadGoods.HeaderText = MapText("เพิ่มรายละเอียดการเบิกอาหาร")
            frmAddEditLoadGoods.ShowMode = SHOW_ADD
            Load frmAddEditLoadGoods
            frmAddEditLoadGoods.Show 1
   
            OKClick = frmAddEditLoadGoods.OKClick
   
            Unload frmAddEditLoadGoods
            Set frmAddEditLoadGoods = Nothing
   
            If OKClick Then
                  GridEX1.ItemCount = CountItem(m_InventoryWHDoc.C_LotItemsWH)
                  GridEX1.Rebind
            End If
   End If

   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdAuto_Click()
   If Trim(txtDocumentNo.Text) = "" And ShowMode = SHOW_ADD Then
      txtDocumentNo.Text = GetDocumentNo(DocumentType)
   End If
End Sub

Private Function GetDocumentNo(DocNoType As Long) As String
Dim No As String
Dim DOC_ID As Long
Dim Cd As CConfigDoc
Dim TempStr As String
Dim I As Long
Dim ServerDateTime As String

   If DocNoType = 2000 Then
      DOC_ID = WH_LOAD_GOODS_BAG
   ElseIf DocNoType = 2001 Then
      DOC_ID = WH_LOAD_GOODS_BULK
   ElseIf DocNoType = 2004 Then
      DOC_ID = WH_LOAD_GOODS_OTHER
   End If
    
    If DOC_ID > 0 Then
       Set Cd = GetObject("CConfigDoc", m_Cd, Trim(str(DOC_ID)), False)
       If Not (Cd Is Nothing) Then
          GetDocumentNo = Cd.GetFieldValue("PREFIX") & Cd.GetFieldValue("CODE1")
          TempStr = ""
          If Cd.GetFieldValue("YEAR_TYPE") = 1 Then
             TempStr = Right(Format(Year(Now) + 543, "0000"), 2)
          ElseIf Cd.GetFieldValue("YEAR_TYPE") = 2 Then
             TempStr = Format(Year(Now) + 543, "0000")
          ElseIf Cd.GetFieldValue("YEAR_TYPE") = 3 Then
             TempStr = Right(Format(Year(Now), "0000"), 2)
          ElseIf Cd.GetFieldValue("YEAR_TYPE") = 4 Then
             TempStr = Format(Year(Now), "0000")
          End If
          GetDocumentNo = GetDocumentNo & TempStr & Cd.GetFieldValue("CODE2")
          TempStr = ""
          If Cd.GetFieldValue("MONTH_TYPE") = 1 Then
             TempStr = Format(Month(Now), "00")
          End If
          GetDocumentNo = GetDocumentNo & TempStr & Cd.GetFieldValue("CODE3")
          TempStr = ""
          For I = 1 To Cd.GetFieldValue("DIGIT_AMOUNT")
             TempStr = TempStr & "0"
          Next I
'          GetDocumentNo = GetDocumentNo & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
'          m_InventoryWHDoc.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
'          m_InventoryWHDoc.CONFIG_DOC_TYPE = DOC_ID
           If Cd.GetFieldValue("AUTO_BEGIN_FLAG") = "Y" Then
               If CheckNewMounth And CheckUniqueNs(INVENTORY_WH_DOC_UNIQUE, GetDocumentNo & Format(1, TempStr), ID) Then
                  GetDocumentNo = GetDocumentNo & Format(1, TempStr) 'เริ่มจาก 1 เสมอ
                  m_InventoryWHDoc.RUNNING_NO = 1
               Else
                  GetDocumentNo = GetDocumentNo & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
                 m_InventoryWHDoc.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
               End If
          Else
               GetDocumentNo = GetDocumentNo & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
                m_InventoryWHDoc.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
          End If
          m_InventoryWHDoc.CONFIG_DOC_TYPE = DOC_ID
       Else
          GetDocumentNo = ""
       End If
    End If
      
End Function
Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
'   If m_HasModify Or ((Not m_HasModify) And (ShowMode = SHOW_ADD)) Then
'      glbErrorLog.LocalErrorMsg = "กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน"
'      glbErrorLog.ShowUserError
'      Exit Sub
'   End If

   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If

   ID1 = GridEX1.Value(10)
   ID2 = GridEX1.Value(2)
   

   If TabStrip1.SelectedItem.Index = 1 Then
      If ID1 <= 0 Then
         m_InventoryWHDoc.C_LotItemsWH.Item(ID2).Flag = "D"
'         m_InventoryWHDoc.C_LotItemsWH.Remove (ID2)
      Else
         m_InventoryWHDoc.C_LotItemsWH.Item(ID2).Flag = "D"
      End If

'      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_InventoryWHDoc.C_LotItemsWH)
      GridEX1.Rebind
      m_HasModify = True
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim OKClick As Boolean
Dim AutoSave As Boolean
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID2 = Val(GridEX1.Value(2))
   OKClick = False
  
   If TabStrip1.SelectedItem.Index = 1 Then
        If DocumentType = 2000 Then
            If Not VerifyAccessRight("INVENTORY-WH_EXPORT" & "_" & DocumentType & "_" & "EDIT" & "_" & "LOAD-BAG", "สามารถโหลดสินค้า BAG ได้") Then
               Call EnableForm(Me, True)
               Exit Sub
            End If
         ElseIf DocumentType = 2001 Then
            If Not VerifyAccessRight("INVENTORY-WH_EXPORT" & "_" & DocumentType & "_" & "EDIT" & "_" & "LOAD-BULK", "สามารถโหลดสินค้า BULK ได้") Then
               Call EnableForm(Me, True)
               Exit Sub
            End If
          ElseIf DocumentType = 2004 Then
            If Not VerifyAccessRight("INVENTORY-WH_EXPORT" & "_" & DocumentType & "_" & "EDIT" & "_" & "LOAD-OTHER", "สามารถโหลดสินค้า อื่นๆ ได้") Then
               Call EnableForm(Me, True)
               Exit Sub
            End If
         End If
         If DocumentType = 2000 Or DocumentType = 2001 Then
            frmAddEditLoadGoods.ID = ID2
            frmAddEditLoadGoods.PART_ITEM_ID = GridEX1.Value(10)
            frmAddEditLoadGoods.PART_TYPE = GridEX1.Value(11)
            frmAddEditLoadGoods.LOCATION_ID = GridEX1.Value(13)
            frmAddEditLoadGoods.DOCUMENT_TYPE = DocumentType
            frmAddEditLoadGoods.DOCUMENT_DATE = uctlDocumentDate.ShowDate
      
            Set frmAddEditLoadGoods.TempLotItemsWH = m_InventoryWHDoc.C_LotItemsWH
            frmAddEditLoadGoods.HeaderText = MapText("แก้ไขรายละเอียดการเบิกอาหาร")
            frmAddEditLoadGoods.ShowMode = SHOW_EDIT
            Load frmAddEditLoadGoods
            frmAddEditLoadGoods.Show 1
            OKClick = frmAddEditLoadGoods.OKClick
            AutoSave = frmAddEditLoadGoods.AutoSave
   
            Unload frmAddEditLoadGoods
            Set frmAddEditLoadGoods = Nothing
            
            If AutoSave Then
               m_HasModify = True
               If Not SaveData Then
                  Exit Sub
               End If
               
               ShowMode = SHOW_EDIT
               ID = m_InventoryWHDoc.INVENTORY_WH_DOC_ID
               m_InventoryWHDoc.QueryFlag = 1
               QueryData (True)
               m_HasModify = False
            End If
   
            If OKClick Then
                  GridEX1.ItemCount = CountItem(m_InventoryWHDoc.C_LotItemsWH)
                  GridEX1.Rebind
            End If
         Else
            
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
      ID = m_InventoryWHDoc.INVENTORY_WH_DOC_ID
      m_InventoryWHDoc.QueryFlag = 1
      QueryData (True)
      m_HasModify = False
      
   ElseIf lMenuChosen = 3 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      OKClick = True
      Unload Me
   End If
    Call TabStrip1_Click
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
   Select Case DocumentType
   Case 2000, 2001, 2004
      lMenuChosen = oMenu.Popup("ใบรายงานการขึ้นอาหารแบบปกติ", "-", "-", "-", "ใบรายงานการขึ้นอาหารแบบมีพื้นหลัง", "-", "ปรับค่าหน้ากระดาษ", "-", "ใบฝากขาย", "-", "ปรับค่าหน้ากระดาษ")
   End Select
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   Call EnableForm(Me, False)

   Set Rc = New CReportConfig
   If DocumentType = 2000 Or DocumentType = 2001 Or DocumentType = 2004 Then
       ReportKey = "CReportLD001"
      If lMenuChosen = 1 Then
         Set Report = New CReportLD001
         
         Picture1.Picture = LoadPicture(glbParameterObj.LoadGoodsPic)
         Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
         Call Report.AddParam(False, "FULL_SHOW")
      
         ReportFlag = True
         Call Report.AddParam(1, "PREVIEW_TYPE")
       ElseIf lMenuChosen = 5 Then
         Set Report = New CReportLD001
         
         Picture1.Picture = LoadPicture(glbParameterObj.LoadGoodsPic)
         Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
         Call Report.AddParam(True, "FULL_SHOW")
      
         ReportFlag = True
         Call Report.AddParam(1, "PREVIEW_TYPE")
      ElseIf lMenuChosen = 9 Then
         Set Report = New CReportNormalRQ1
         
'         Picture1.Picture = LoadPicture(glbParameterObj.LoadGoodsPic)
'         Call Report.AddParam(Picture1.Picture, "BACK_GROUND")
'         Call Report.AddParam(True, "FULL_SHOW")

        Call Report.AddParam(DOCUMENT_ID_RQ, "DOCUMENT_ID_RQ")
      
         ReportFlag = True
         Call Report.AddParam(1, "PREVIEW_TYPE")
      End If
   End If

   If Not Report Is Nothing Then

      Call Report.AddParam(m_InventoryWHDoc.INVENTORY_WH_DOC_ID, "INVENTORY_WH_DOC_ID")
      Call Report.AddParam(lMenuChosen, "REPORT_TYPE")
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

   Else
      If lMenuChosen = 7 Then
         ReportMode = 2
      End If
         
         Set Rc = New CReportConfig
         Rc.REPORT_KEY = ReportKey
         Call Rc.QueryData(m_Rs, iCount)
         
         If Not m_Rs.EOF Then
            Call Rc.PopulateFromRS(1, m_Rs)
            frmReportConfig.ShowMode = SHOW_EDIT
            frmReportConfig.ID = Rc.REPORT_CONFIG_ID
         Else
            frmReportConfig.ShowMode = SHOW_ADD
         End If
         
         frmReportConfig.ReportMode = ReportMode
         frmReportConfig.ReportKey = ReportKey
         frmReportConfig.HeaderText = HeaderText
         Load frmReportConfig
         frmReportConfig.Show 1
         
         Unload frmReportConfig
         Set frmReportConfig = Nothing
         
         Set Rc = Nothing
   End If

   Call EnableForm(Me, True)
End Sub

Private Sub cmdLoadWeight_Click()
Dim CW As CWeight
   If Not VerifyDate(lblDocumentDate, uctlDocumentDate, False) Then
      Exit Sub
   End If

   TempWeigth = txtEntryWeight.Text

'   frmAddWeight.SupplierCode = uctlCustomerLookup.MyTextBox.Text 'SupplierCode
   frmAddWeight.TruckNo = Trim(txtTruckNo.Text)
   frmAddWeight.DocumentDate = uctlDocumentDate.ShowDate
   frmAddWeight.ShowMode = 1
   frmAddWeight.HeaderText = MapText("เพิ่มรายการน้ำหนักจากโปรแกรมเครื่องชั่ง")
   Load frmAddWeight
   frmAddWeight.Show 1
   OKClick = frmAddWeight.OKClick
   CancelWeigth = frmAddWeight.CancelWeigth
   NotCheckWeigth = frmAddWeight.TempWeigth
   EditWeigth = frmAddWeight.EditTempWeigth
   If EditWeigth Then
      TempWeigth = iniNotCheckWeigth
   End If
   If OKClick Then
      Set TempWeight = frmAddWeight.TempCollection
      Set CW = GetObject("CWeight", TempWeight, "1")
      txtWeightNo.Text = CW.WEIGHT_ID
      txtEntryWeight.Text = CW.WEIGHT1
      txtExitWeight.Text = CW.WEIGHT2
      txtWeightAmount.Text = CW.NetWeight
      txtWeightNote.Text = CW.REMARK
      txtEntryWeightDate.Text = CW.DateShow1
      txtExitWeightDate.Text = CW.DateShow2
      txtEntryWeightTime.Text = CW.Time1
      txtExitWeightTime.Text = CW.Time2
   End If
  If CancelWeigth Then
      txtEntryWeight.Text = ""
      txtExitWeight.Text = ""
      txtWeightAmount.Text = ""
      txtWeightNote.Text = ""
      txtWeightNo.Text = ""
      txtEntryWeightDate.Text = ""
      txtExitWeightDate.Text = ""
      txtEntryWeightTime.Text = ""
      txtExitWeightTime.Text = ""
   ElseIf NotCheckWeigth Then
      txtWeightNo.Text = "TW001"
      txtEntryWeight.Text = iniNotCheckWeigth
      txtExitWeight.Text = iniNotCheckWeigth
      txtWeightAmount.Text = iniNotCheckWeigth
      txtWeightNote.Text = "ออกเอกสารก่อนรถมารับจริง"
      txtEntryWeightDate.Text = ""
      txtExitWeightDate.Text = ""
      txtEntryWeightTime.Text = ""
      txtExitWeightTime.Text = ""
   End If
   Unload frmAddWeight
   Set frmAddWeight = Nothing
End Sub



Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
      Set uctlCustomerLookup.MyCollection = m_Customers
      
'      Call LoadLocation(uctlPlaceLookup.MyCombo, m_Locations, 2)
'      Set uctlPlaceLookup.MyCollection = m_Locations
      
      Call LoadEmployee(uctlEmpCheckCar.MyCombo, m_Employees, 11) '11 คือเลือกเฉพาะเจ้าหน้าที่โกดัง
      Set uctlEmpCheckCar.MyCollection = m_Employees
      Call LoadEmployee(uctlEmpCheckProduct.MyCombo, m_Employees2, 11) '11 คือเลือกเฉพาะเจ้าหน้าที่โกดัง
      Set uctlEmpCheckProduct.MyCollection = m_Employees2
      
      Call LoadConfigDoc(Nothing, m_Cd)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_InventoryWHDoc.QueryFlag = 1 'เอาลูกหลานด้วย
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         uctlDocumentDate.ShowDate = Now
         m_InventoryWHDoc.QueryFlag = 0
         Call QueryData(False)
         uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, CustomerID)
         txtDoNo.Text = DOCUMENT_NO
         txtTruckNo.Text = TRUCK_NO
      End If
      iniNotCheckWeigth = "9999999"

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

   SSFrame1.Top = 0
   SSFrame1.Width = ScaleWidth
   SSFrame1.HEIGHT = ScaleHeight
   
   pnlHeader.Width = ScaleWidth
   
   GridEX1.Width = ScaleWidth - 300
   GridEX1.HEIGHT = ScaleHeight - GridEX1.Top - 640
   
   TabStrip1.Width = GridEX1.Width
'   SSFrame4.Top = GridEX1.Top
'   SSFrame4.Width = GridEX1.Width
'   SSFrame4.HEIGHT = GridEX1.HEIGHT
   
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
  cmdPrint.Top = ScaleHeight - 580
  cmdOther.Top = ScaleHeight - 580
  
  
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
    cmdPrint.Left = cmdOK.Left - cmdPrint.Width - 50
    cmdOther.Left = cmdPrint.Left - cmdOther.Width - 50
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_InventoryWHDoc = Nothing
   Set m_Customers = Nothing
   Set m_Locations = Nothing
   Set m_Employees = Nothing
   Set m_Employees2 = Nothing
   Set m_Cd = Nothing
   Set TempWeight = Nothing
   Set m_Weight = Nothing
   Set m_CollLotDoc = Nothing
   Set m_CollJob = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn
Dim I As Long

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.NAME = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 10
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 600
   Col.Caption = "ลำดับ"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 3500
   Col.Caption = MapText("เบอร์อาหาร")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 3500
   Col.Caption = MapText("รายละเอียด")
   
   If DocumentType = 2000 Then
      Set Col = GridEX1.Columns.add '5
      Col.Width = 1100
      Col.Caption = MapText("นน./ถุง")
      
      Set Col = GridEX1.Columns.add '6
      Col.Width = 1300
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวนถุง")
   Else
      Set Col = GridEX1.Columns.add '5
      Col.Width = 0
      Col.Caption = MapText("นน./ถุง")
      
      Set Col = GridEX1.Columns.add '6
      Col.Width = 0
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวนถุง")
   End If

   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 1300
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("น้ำหนักรวม")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("น้ำหนักขึ้นจริง")
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 3500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("หมายเลข SO")
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 0
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("PART_ITEM_ID")
   
   Set Col = GridEX1.Columns.add '11
   Col.Width = 0
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("PART_TYPE")
   
   Set Col = GridEX1.Columns.add '12
   Col.Width = 2000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("สถานะการขึ้นอาหาร")
   
   Set Col = GridEX1.Columns.add '13
   Col.Width = 0
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("LOCATION_ID")

End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
'   SSFrame4.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่ใบขึ้นอาหาร"))
   Call InitNormalLabel(lblTruckNo, MapText("ทะเบียนรถ"))
   Call InitNormalLabel(lblDesc, MapText("หมายเหตุ"))
   Call InitNormalLabel(lblDoNo, MapText("IVD ชั่วคราว"))
'   Call InitNormalLabel(lblSoNo, MapText("เลขที่ SO"))
   
   Call InitNormalLabel(lblConLog, MapText("สภาพรถขนส่ง"))
   Call InitNormalLabel(lblCon1, MapText("การคลุมผ้าใบรถขนส่ง"))
   Call InitNormalLabel(lblCon2, MapText("ความสะอาดรถ"))
   Call InitNormalLabel(lblCon3, MapText("ความพร้อมของกะบะและพื้นรถ"))
   Call InitNormalLabel(lblCheckCar, MapText("ผู้ตรวจสภาพรถ"))
   Call InitNormalLabel(lblCheckProduct, MapText("พนักงานเช็คสินค้า"))
   
   
   Call InitNormalLabel(lblWeightIn, MapText("น้ำหนักเข้า"))
   Call InitNormalLabel(lblWeightOut, MapText("น้ำหนักออก"))
   Call InitNormalLabel(lblWeightTotal, MapText("น้ำหนักสุทธิ"))
   Call InitNormalLabel(lblNote, MapText("หมายเหตุ"))
   Call InitNormalLabel(Label1, MapText("กก."))
   Call InitNormalLabel(Label2, MapText("กก."))
   Call InitNormalLabel(Label4, MapText("กก."))
   Call InitNormalLabel(lblWeightNo, MapText("เลขที่ใบชั่ง"))
   Call InitNormalLabel(lblWeightTime, MapText("เวลาชั่ง"))
   Call InitCheckBox(sscConsignment, "ฝากขาย")
   
   If Not VerifyAccessRight("INVENTORY-WH_EXPORT" & "_" & DocumentType & "_" & "EDIT" & "_" & "CONSIGNMENT", "สามารถออกใบฝากขายได้", 2) Then
      sscConsignment.Enabled = False
   Else
     sscConsignment.Enabled = True
   End If
   
   uctlCustomerLookup.MyTextBox.SetKeySearch ("CUSTOMER_CODE")
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่เอกสาร"))
   Call InitNormalLabel(lblCustomerNo, MapText("รหัสลูกค้า"))
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
'   txtDocumentNo.Enabled = False
   Call txtDoNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
'   Call txtSoNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtTruckNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtEntryWeight.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtEntryWeight.Enabled = False
   Call txtExitWeight.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtExitWeight.Enabled = False
   Call txtWeightAmount.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtWeightAmount.Enabled = False
   Call txtWeightNote.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtWeightNote.Enabled = False
   Call txtWeightNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtWeightNo.Enabled = False
   Call txtEntryWeightDate.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtEntryWeightDate.Enabled = False
   Call txtExitWeightDate.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtExitWeightDate.Enabled = False
   Call txtEntryWeightTime.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtEntryWeightTime.Enabled = False
   Call txtExitWeightTime.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtExitWeightTime.Enabled = False
   
   Call InitCombo(cboCon1)
   Call InitCombo(cboCon2)
   Call InitCombo(cboCon3)
   
   Call InitCboCondition(cboCon1, 1)
   Call InitCboCondition(cboCon2, 2)
   Call InitCboCondition(cboCon3, 3)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
'   SSFrame2.Enabled = False
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOther.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdLoadWeight.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdOther, MapText("อื่นๆ"))
   Call InitMainButton(cmdLoadWeight, MapText("โหลดน้ำหนัก"))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   Dim str As String
   Select Case DocumentType
    Case 2000, 2001, 2004
          str = "รายการสินค้า"
   End Select
   TabStrip1.Tabs.add().Caption = MapText(str)
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   Call glbDaily.RollbackTransaction
   Call EnableForm(Me, True)
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
'   OKClick = False
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_InventoryWHDoc = New CInventoryWHDoc
   Set m_Weight = New CWeight
   Set m_Customers = New Collection
   Set m_Locations = New Collection
   Set m_Employees = New Collection
   Set m_Employees2 = New Collection
   Set m_Cd = New Collection
   Set TempWeight = New Collection
   Set m_CollLotDoc = New Collection
   Set m_CollJob = New Collection
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   If TabStrip1.SelectedItem.Index = 5 Then
      RowBuffer.RowStyle = RowBuffer.Value(7)
   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim I As Long
Dim CountItem As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_InventoryWHDoc.C_LotItemsWH Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If
      
      Dim CR As CLotItemWH
      Dim LTD As CLotDoc
      Dim PD As CPalletDoc
      If m_InventoryWHDoc.C_LotItemsWH.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_InventoryWHDoc.C_LotItemsWH, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
      
      Values(1) = CR.LOT_ITEM_WH_ID
      Values(2) = RealIndex
      If CR.PART_ITEM_ID = -1 Then
         Values(3) = CR.FEATURE_CODE
         Values(4) = CR.FEATURE_DESC
      Else
         Values(3) = CR.PART_NO
         Values(4) = CR.PART_DESC
      End If
      If DocumentType = 2001 Then
         Values(5) = ""
         Values(6) = ""
      Else
         Values(5) = CR.WEIGHT_PER_PACK
         Values(6) = CR.PACK_AMOUNT
      End If
      Values(7) = CR.TX_AMOUNT
      Values(8) = CR.LOAD_TRUE
      Values(9) = CR.DOCUMENT_NO_SO
      Values(10) = CR.PART_ITEM_ID
      Values(11) = CR.PART_TYPE_NO
      If CR.LOAD_TRUE <> CR.TX_AMOUNT Or CR.LOAD_TRUE <= 0 Then
         Values(12) = "ยังไม่สมบูรณ์"
         CR.LOAD_AMOUNT_FLAG = "N" 'เปลี่ยนเป็น รอขึ้นอาหาร
      Else
         Values(12) = "สมบูรณ์"
      End If
      Values(13) = CR.LOCATION_ID

   End If
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub sscConsignment_Click(Value As Integer)
m_HasModify = True
End Sub

'Private Sub SSFrame2_Click()
'   If Not VerifyAccessRight("INVENTORY-WH_EXPORT" & "_" & DocumentType & "_" & "EDIT" & "_" & "CONSIGNMENT", "สามารถออกใบฝากขายได้") Then
'      Call EnableForm(Me, True)
'      Exit Sub
'   End If
'
'   SSFrame2.Enabled = True
'   uctlPlaceLookup.MyCombo.ListIndex = IDToListIndex(uctlPlaceLookup.MyCombo, 224)
'End Sub

Private Sub TabStrip1_Click()
   GridEX1.Visible = False
'   SSFrame4.Visible = False
   If TabStrip1.SelectedItem.Index = 1 Then
     Call EnableDisableButton(True)
      GridEX1.Visible = True
      Call InitGrid1
      GridEX1.ItemCount = CountItem(m_InventoryWHDoc.C_LotItemsWH)
      GridEX1.Rebind
   End If
End Sub



Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtDoNo_Change()
   m_HasModify = True
End Sub

Private Sub txtDeliveryNo_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub txtDocumentNo_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtDueAmount_Change()
   m_HasModify = True
End Sub


Private Sub txtPrNo_Change()
   m_HasModify = True
End Sub

Private Sub txtQueNo_Change()
   m_HasModify = True
End Sub

Private Sub txtReceiver_Change()
   m_HasModify = True
End Sub

Private Sub txtSender_Change()
   m_HasModify = True
End Sub

Private Sub txtTotal_Change()
   m_HasModify = True
End Sub

Private Sub txtEntryWeight_Change()
   m_HasModify = True
End Sub

Private Sub txtEntryWeightDate_Change()
   m_HasModify = True
End Sub

Private Sub txtEntryWeightTime_Change()
   m_HasModify = True
End Sub

Private Sub txtExitWeight_Change()
   m_HasModify = True
End Sub



Private Sub txtExitWeightDate_Change()
   m_HasModify = True
End Sub

Private Sub txtExitWeightTime_Change()
   m_HasModify = True
End Sub

Private Sub txtTruckNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlSetupDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlDeliveryLookup_Change()
   m_HasModify = True
End Sub

Private Sub txtWeightAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtWeightNo_Change()
   m_HasModify = True
End Sub

Private Sub txtWeightNote_Change()
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlDueDate_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlEntryTime_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlExitTime_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlCustomerLookup_Change()
Dim Ct As CCustomer
Dim CustomerID As Long
   If uctlCustomerLookup.MyCombo.ListIndex < 0 Then
      Exit Sub
   End If
   
   CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   If CustomerID > 0 Then
      Set Ct = GetCustomer(m_Customers, Trim(str(CustomerID)))
     LocationIm = Ct.LOCATION_ID
   Else
      LocationIm = -1
   End If
   
   m_HasModify = True
End Sub
Private Sub PopulateGuiID(BD As CBillingDoc)
Dim Di As CSupItem

   For Each Di In BD.SupItems
      If Di.Flag = "A" Then
         Di.LINK_ID = GetNextGuiID(BD)
      End If
   Next Di
End Sub

Private Function GetNextGuiID(BD As CBillingDoc) As Long
Dim Di As CSupItem
Dim MaxId As Long

   MaxId = 0
   For Each Di In BD.SupItems
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   GetNextGuiID = MaxId + 1
End Function
Private Sub EnableDisableButton(En As Boolean)
   If En Then
      If ShowMode = SHOW_EDIT Then
         cmdAdd.Enabled = True '(m_InventoryWHDoc.OLD_COMMIT_FLAG = "N")
         cmdEdit.Enabled = True '(m_InventoryWHDoc.COMMIT_FLAG = "N")
         cmdDelete.Enabled = True '(m_InventoryWHDoc.OLD_COMMIT_FLAG = "N")
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

Private Sub uctlEmpCheckCar_Change()
m_HasModify = True
End Sub

Private Sub uctlEmpCheckProduct_Change()
m_HasModify = True
End Sub

