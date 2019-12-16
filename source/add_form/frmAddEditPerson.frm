VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddEditPerson 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditPerson.frx":0000
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
      Height          =   8535
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   11700
         _ExtentX        =   20638
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
      Begin Threed.SSFrame fraPerson 
         Height          =   5475
         Left            =   120
         TabIndex        =   44
         Top             =   2250
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   9657
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.ComboBox cboPosition 
            BeginProperty Font 
               Name            =   "AngsanaUPC"
               Size            =   9
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   540
            Width           =   3135
         End
         Begin VB.ComboBox cboWorkStatus 
            BeginProperty Font 
               Name            =   "AngsanaUPC"
               Size            =   9
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   100
            Width           =   3135
         End
         Begin VB.ComboBox cboCauseOut 
            BeginProperty Font 
               Name            =   "AngsanaUPC"
               Size            =   9
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1440
            Width           =   3135
         End
         Begin VB.ComboBox cboBloodGroup 
            BeginProperty Font 
               Name            =   "AngsanaUPC"
               Size            =   9
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   9480
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1920
            Width           =   1365
         End
         Begin VB.ComboBox cbotahan 
            BeginProperty Font 
               Name            =   "AngsanaUPC"
               Size            =   9
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   3360
            Width           =   2685
         End
         Begin VB.ComboBox cboSadsana 
            BeginProperty Font 
               Name            =   "AngsanaUPC"
               Size            =   9
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   2400
            Width           =   2685
         End
         Begin VB.ComboBox cboMary 
            BeginProperty Font 
               Name            =   "AngsanaUPC"
               Size            =   9
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   2880
            Width           =   2685
         End
         Begin VB.ComboBox cboBank 
            BeginProperty Font 
               Name            =   "AngsanaUPC"
               Size            =   9
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   2400
            Width           =   3135
         End
         Begin prjFarmManagement.uctlDate uctlInDate 
            Height          =   405
            Left            =   1440
            TabIndex        =   7
            Top             =   540
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   714
         End
         Begin prjFarmManagement.uctlTextBox txtBranchBank 
            Height          =   435
            Left            =   7200
            TabIndex        =   22
            Top             =   2880
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlDate uctlBirthDate 
            Height          =   405
            Left            =   1440
            TabIndex        =   6
            Top             =   120
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   714
         End
         Begin prjFarmManagement.uctlDate uctlPassDate 
            Height          =   405
            Left            =   1440
            TabIndex        =   8
            Top             =   960
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   714
         End
         Begin prjFarmManagement.uctlDate uctlOutDate 
            Height          =   405
            Left            =   1440
            TabIndex        =   9
            Top             =   1380
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   714
         End
         Begin prjFarmManagement.uctlTextBox txtNickName 
            Height          =   435
            Left            =   8280
            TabIndex        =   27
            Top             =   4440
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtWeight 
            Height          =   435
            Left            =   1800
            TabIndex        =   15
            Top             =   1920
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtNation 
            Height          =   435
            Left            =   1800
            TabIndex        =   24
            Top             =   3960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtDepositNo 
            Height          =   435
            Left            =   7200
            TabIndex        =   23
            Top             =   3360
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtRok 
            Height          =   435
            Left            =   1800
            TabIndex        =   26
            Top             =   4440
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtHight 
            Height          =   435
            Left            =   5280
            TabIndex        =   16
            Top             =   1920
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtInterNational 
            Height          =   435
            Left            =   5280
            TabIndex        =   25
            Top             =   3960
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtSalary 
            Height          =   435
            Left            =   7200
            TabIndex        =   12
            Top             =   975
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtOt 
            Height          =   435
            Left            =   9960
            TabIndex        =   13
            Top             =   975
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtHomeNo 
            Height          =   435
            Left            =   1800
            TabIndex        =   28
            Top             =   4920
            Width           =   1815
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtMobileNo 
            Height          =   435
            Left            =   5280
            TabIndex        =   29
            Top             =   4920
            Width           =   1575
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtEmail 
            Height          =   435
            Left            =   8280
            TabIndex        =   30
            Top             =   4920
            Width           =   3375
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin VB.Label lblBirthDate 
            Alignment       =   1  'Right Justify
            Caption         =   "lblBirthDate"
            Height          =   315
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblWorkStatus 
            Alignment       =   1  'Right Justify
            Caption         =   "lblWorkStatus"
            Height          =   255
            Left            =   5400
            TabIndex        =   72
            Top             =   240
            Width           =   1725
         End
         Begin VB.Label lblInDate 
            Alignment       =   1  'Right Justify
            Caption         =   "lblInDate"
            Height          =   315
            Left            =   120
            TabIndex        =   71
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblPassDate 
            Alignment       =   1  'Right Justify
            Caption         =   "lblPassDate"
            Height          =   315
            Left            =   120
            TabIndex        =   70
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lblOutDate 
            Alignment       =   1  'Right Justify
            Caption         =   "lblOutDate"
            Height          =   315
            Left            =   240
            TabIndex        =   69
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label lblNickName 
            Alignment       =   1  'Right Justify
            Caption         =   "lblNickName"
            Height          =   255
            Left            =   7320
            TabIndex        =   68
            Top             =   4560
            Width           =   885
         End
         Begin VB.Label lblPosition 
            Alignment       =   1  'Right Justify
            Caption         =   "lblPosition"
            Height          =   255
            Left            =   5400
            TabIndex        =   67
            Top             =   600
            Width           =   1725
         End
         Begin VB.Label lblSalaryUnit 
            Alignment       =   1  'Right Justify
            Caption         =   "lblSalaryUnit"
            Height          =   255
            Left            =   8880
            TabIndex        =   66
            Top             =   1080
            Width           =   435
         End
         Begin VB.Label lblCauseOut 
            Alignment       =   1  'Right Justify
            Caption         =   "lblCauseOut"
            Height          =   375
            Left            =   5640
            TabIndex        =   65
            Top             =   1560
            Width           =   1485
         End
         Begin VB.Label lblWeight 
            Alignment       =   1  'Right Justify
            Caption         =   "lblWeight"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label lblWeightUnit 
            Alignment       =   1  'Right Justify
            Caption         =   "lblWeightUnit"
            Height          =   255
            Left            =   3600
            TabIndex        =   63
            Top             =   2040
            Width           =   885
         End
         Begin VB.Label lblHight 
            Alignment       =   1  'Right Justify
            Caption         =   "lblHight"
            Height          =   255
            Left            =   4320
            TabIndex        =   62
            Top             =   2040
            Width           =   885
         End
         Begin VB.Label lblHightUnit 
            Alignment       =   1  'Right Justify
            Caption         =   "lblHightUnit"
            Height          =   255
            Left            =   6840
            TabIndex        =   61
            Top             =   2040
            Width           =   1125
         End
         Begin VB.Label lblBloodGroup 
            Alignment       =   1  'Right Justify
            Caption         =   "lblBloodGroup"
            Height          =   375
            Left            =   8400
            TabIndex        =   60
            Top             =   2040
            Width           =   1005
         End
         Begin VB.Label lblNation 
            Alignment       =   1  'Right Justify
            Caption         =   "lblNation"
            Height          =   495
            Left            =   240
            TabIndex        =   59
            Top             =   4080
            Width           =   1485
         End
         Begin VB.Label lblInterNational 
            Alignment       =   1  'Right Justify
            Caption         =   "lblInterNational"
            Height          =   255
            Left            =   3720
            TabIndex        =   58
            Top             =   4080
            Width           =   1485
         End
         Begin VB.Label lblSalary 
            Alignment       =   1  'Right Justify
            Caption         =   "lblSalary"
            Height          =   255
            Left            =   6120
            TabIndex        =   57
            Top             =   1080
            Width           =   1005
         End
         Begin VB.Label lblOt 
            Alignment       =   1  'Right Justify
            Caption         =   "lblOt"
            Height          =   255
            Left            =   9360
            TabIndex        =   56
            Top             =   1080
            Width           =   525
         End
         Begin VB.Label lblOtUnit 
            Alignment       =   1  'Right Justify
            Caption         =   "lblOtUnit"
            Height          =   255
            Left            =   10800
            TabIndex        =   55
            Top             =   1080
            Width           =   885
         End
         Begin VB.Label lblSadsana 
            Alignment       =   1  'Right Justify
            Caption         =   "lblSadsana"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   2520
            Width           =   1485
         End
         Begin VB.Label lbltahan 
            Alignment       =   1  'Right Justify
            Caption         =   "lbltahan"
            Height          =   255
            Left            =   0
            TabIndex        =   53
            Top             =   3480
            Width           =   1635
         End
         Begin VB.Label lblMary 
            Alignment       =   1  'Right Justify
            Caption         =   "lblMary"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   3000
            Width           =   1485
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            Caption         =   "lblBank"
            Height          =   255
            Left            =   5640
            TabIndex        =   51
            Top             =   2520
            Width           =   1485
         End
         Begin VB.Label lblBranchBank 
            Alignment       =   1  'Right Justify
            Caption         =   "lblBranchBank"
            Height          =   255
            Left            =   5640
            TabIndex        =   50
            Top             =   3000
            Width           =   1485
         End
         Begin VB.Label lblDepositNo 
            Alignment       =   1  'Right Justify
            Caption         =   "lblDepositNo"
            Height          =   375
            Left            =   5640
            TabIndex        =   49
            Top             =   3480
            Width           =   1485
         End
         Begin VB.Label lblMobileNo 
            Alignment       =   1  'Right Justify
            Caption         =   "lblMobileNo"
            Height          =   255
            Left            =   3720
            TabIndex        =   48
            Top             =   5040
            Width           =   1485
         End
         Begin VB.Label lblRok 
            Alignment       =   1  'Right Justify
            Caption         =   "lblRok"
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   4560
            Width           =   1485
         End
         Begin VB.Label lblHomeNo 
            Alignment       =   1  'Right Justify
            Caption         =   "lblHomeNo"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   5040
            Width           =   1605
         End
         Begin VB.Label lblEmail 
            Alignment       =   1  'Right Justify
            Caption         =   "lblEmail"
            Height          =   255
            Left            =   6960
            TabIndex        =   45
            Top             =   5040
            Width           =   1125
         End
      End
      Begin VB.ComboBox cboSex 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1320
         Width           =   1485
      End
      Begin prjFarmManagement.uctlTextBox txtPersonName 
         Height          =   435
         Left            =   2280
         TabIndex        =   0
         Top             =   840
         Width           =   2775
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   12240
         _ExtentX        =   21590
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPersonLname 
         Height          =   435
         Left            =   7440
         TabIndex        =   1
         Top             =   840
         Width           =   3375
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtPersonCode 
         Height          =   435
         Left            =   8760
         TabIndex        =   4
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   767
      End
      Begin Threed.SSFrame fraGrid 
         Height          =   5400
         Left            =   120
         TabIndex        =   42
         Top             =   2280
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   9525
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin GridEX20.GridEX GridEX1 
            Height          =   5400
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   11700
            _ExtentX        =   20638
            _ExtentY        =   9525
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
            Column(1)       =   "frmAddEditPerson.frx":27A2
            Column(2)       =   "frmAddEditPerson.frx":286A
            FormatStylesCount=   5
            FormatStyle(1)  =   "frmAddEditPerson.frx":290E
            FormatStyle(2)  =   "frmAddEditPerson.frx":2A6A
            FormatStyle(3)  =   "frmAddEditPerson.frx":2B1A
            FormatStyle(4)  =   "frmAddEditPerson.frx":2BCE
            FormatStyle(5)  =   "frmAddEditPerson.frx":2CA6
            ImageCount      =   0
            PrinterProperties=   "frmAddEditPerson.frx":2D5E
         End
      End
      Begin VB.Label lblPersonCode 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPersonCode"
         Height          =   315
         Left            =   6960
         TabIndex        =   41
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblPersonSex 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPersonSex"
         Height          =   315
         Left            =   120
         TabIndex        =   40
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lblPersonLname 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPersonLname"
         Height          =   315
         Left            =   5880
         TabIndex        =   39
         Top             =   960
         Width           =   1455
      End
      Begin Threed.SSCheck chkPersonOut 
         Height          =   435
         Left            =   3840
         TabIndex        =   3
         Top             =   1320
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8520
         TabIndex        =   34
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPerson.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10200
         TabIndex        =   35
         Top             =   7800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPerson.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   33
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPerson.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblPersonName 
         Alignment       =   1  'Right Justify
         Caption         =   "lblPersonName"
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   960
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmAddEditPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Employee As CEmployee
Private m_Employees As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public id As Long

Private FileName As String
Private m_SumUnit As Double
Public TempCollection As Collection
Public TempCollection2 As Collection
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
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

   chkPersonOut.Value = FlagToCheck(m_Employee.EXTERNAL_FLAG)
  txtPersonCode.Text = m_Employee.EMP_CODE
  txtPersonName.Text = m_Employee.NAME
  txtPersonLname.Text = m_Employee.LASTNAME
   txtNickName.Text = m_Employee.EName.NICK_NAME
cboPosition.ListIndex = IDToListIndex(cboPosition, m_Employee.CURRENT_POSITION)
   
   uctlBirthDate.ShowDate = m_Employee.BIRTH_DATE
   uctlInDate.ShowDate = m_Employee.ENTRY_DATE
 uctlPassDate.ShowDate = m_Employee.PASS_DATE
  uctlOutDate.ShowDate = m_Employee.RESIGN_DATE
   
cboSadsana.ListIndex = IDToListIndex(cboSadsana, m_Employee.RELIGIOUS_ID)
cboMary.ListIndex = IDToListIndex(cboMary, m_Employee.MARITAL_ID)
cbotahan.ListIndex = IDToListIndex(cbotahan, m_Employee.MILITALY_ID)
cboWorkStatus.ListIndex = IDToListIndex(cboWorkStatus, m_Employee.WORK_STATUS_ID)
cboBloodGroup.ListIndex = IDToListIndex(cboBloodGroup, m_Employee.BLOOD_GROUP)
cboCauseOut.ListIndex = IDToListIndex(cboCauseOut, m_Employee.RESIGN_REASON)
cboBank.ListIndex = IDToListIndex(cboBank, m_Employee.BANK_ID)
cboSex.ListIndex = IDToListIndex(cboSex, m_Employee.SEX_ID)
   
  txtWeight.Text = m_Employee.WEIGHT
   txtHight.Text = m_Employee.HEIGHT
  txtRok.Text = m_Employee.DISCEASE
txtMobileNo.Text = m_Employee.MOBILE_PHONE
    txtEmail.Text = m_Employee.EMAIL_ADDRESS
txtSalary.Text = m_Employee.CURRENT_SALARY
   txtBranchBank.Text = m_Employee.BANK_BRANCH
  txtDepositNo.Text = m_Employee.BANK_ACCOUNT
    txtOt.Text = m_Employee.OT_RATE
      txtNation.Text = m_Employee.NATIONALITY
   txtInterNational.Text = m_Employee.RACE
    txtHomeNo.Text = m_Employee.HOME_PHONE

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

   If Not VerifyTextControl(lblPersonName, txtPersonName, False) Then
      Exit Function
   End If
If Not VerifyTextControl(lblPersonLname, txtPersonLname, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblPersonCode, txtPersonCode, False) Then
      Exit Function
   End If
If Not VerifyTextControl(lblSalary, txtSalary, False) Then
      Exit Function
   End If
'If Not VerifyTextControl(lblBranchBank, txtBranchBank, False) Then
 '     Exit Function
  ' End If
   'If Not VerifyTextControl(lblDepositNo, txtDepositNo, False) Then
    '  Exit Function
   'End If

   'If Not VerifyDate(lblBirthDate, uctlBirthDate, False) Then
    '  Exit Function
   'End If
   
   'If Not VerifyDate(lblInDate, uctlInDate, False) Then
    '  Exit Function
   'End If
   
   If Not VerifyCombo(lblPersonSex, cboSex, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPosition, cboPosition, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblWorkStatus, cboWorkStatus, False) Then
      Exit Function
   End If
   'If Not VerifyCombo(lblBank, cboBank, False) Then
    '  Exit Function
   'End If
   If Not CheckUniqueNs(PERSON_CODE, txtPersonCode.Text, id) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtPersonCode.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   m_Employee.EMP_ID = id
   m_Employee.AddEditMode = ShowMode
   m_Employee.EXTERNAL_FLAG = Check2Flag(chkPersonOut.Value)
   m_Employee.EMP_CODE = txtPersonCode.Text
   m_Employee.NAME = txtPersonName.Text
   m_Employee.LASTNAME = txtPersonLname.Text
   m_Employee.CURRENT_POSITION = cboPosition.ItemData(Minus2Zero(cboPosition.ListIndex))
   m_Employee.PASS_STATUS = "Y"
   m_Employee.BIRTH_DATE = uctlBirthDate.ShowDate
   m_Employee.ENTRY_DATE = uctlInDate.ShowDate
   m_Employee.PASS_DATE = uctlPassDate.ShowDate
   m_Employee.RESIGN_DATE = uctlOutDate.ShowDate
   
m_Employee.RELIGIOUS_ID = cboSadsana.ItemData(Minus2Zero(cboSadsana.ListIndex))
m_Employee.MARITAL_ID = cboMary.ItemData(Minus2Zero(cboMary.ListIndex))
m_Employee.MILITALY_ID = cbotahan.ItemData(Minus2Zero(cbotahan.ListIndex))
m_Employee.WORK_STATUS_ID = cboWorkStatus.ItemData(Minus2Zero(cboWorkStatus.ListIndex))
m_Employee.BLOOD_GROUP = cboBloodGroup.ItemData(Minus2Zero(cboBloodGroup.ListIndex))
m_Employee.RESIGN_REASON = cboCauseOut.ItemData(Minus2Zero(cboCauseOut.ListIndex))
m_Employee.BANK_ID = cboBank.ItemData(Minus2Zero(cboBank.ListIndex))
m_Employee.SEX_ID = cboSex.ItemData(Minus2Zero(cboSex.ListIndex))

   m_Employee.WEIGHT = Val(txtWeight.Text)
   m_Employee.HEIGHT = Val(txtHight.Text)
   m_Employee.DISCEASE = txtRok.Text
   m_Employee.MOBILE_PHONE = txtMobileNo.Text
   m_Employee.EMAIL_ADDRESS = txtEmail.Text
   m_Employee.CURRENT_SALARY = Val(txtSalary.Text)
   m_Employee.BANK_BRANCH = txtBranchBank.Text
   m_Employee.BANK_ACCOUNT = txtDepositNo.Text
   m_Employee.OT_RATE = Val(txtOt.Text)
      m_Employee.NATIONALITY = txtNation.Text
   m_Employee.RACE = txtInterNational.Text
   m_Employee.HOME_PHONE = txtHomeNo.Text

   '
   
   m_Employee.EmpName.AddEditMode = ShowMode
   m_Employee.EName.AddEditMode = ShowMode
   m_Employee.EName.LONG_NAME = txtPersonName.Text
   m_Employee.EName.LAST_NAME = txtPersonLname.Text
    m_Employee.EName.NICK_NAME = txtNickName.Text

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

Private Sub cboBank_Change()
m_HasModify = True
End Sub

Private Sub cboBank_Click()
m_HasModify = True
End Sub

Private Sub cboBloodGroup_Change()
m_HasModify = True
End Sub

Private Sub cboBloodGroup_Click()
m_HasModify = True
End Sub

Private Sub cboCauseOut_Change()
m_HasModify = True
End Sub

Private Sub cboCauseOut_Click()
m_HasModify = True
End Sub
Private Sub cboMary_Change()
m_HasModify = True
End Sub

Private Sub cboMary_Click()
m_HasModify = True
End Sub

Private Sub cboPosition_Change()
m_HasModify = True
End Sub

Private Sub cboPosition_Click()
m_HasModify = True
End Sub

Private Sub cboSadsana_Change()
m_HasModify = True
End Sub

Private Sub cboSadsana_Click()
m_HasModify = True
End Sub

Private Sub cboSex_Change()
m_HasModify = True
End Sub

Private Sub cboSex_Click()
m_HasModify = True
End Sub

Private Sub cbotahan_Change()
m_HasModify = True
End Sub

Private Sub cbotahan_Click()
m_HasModify = True
End Sub

Private Sub cboWorkStatus_Change()
m_HasModify = True
End Sub

Private Sub cboWorkStatus_Click()
m_HasModify = True
End Sub

Private Sub chkPersonOut_Click(Value As Integer)
m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If

   OKClick = False
   If TabStrip1.SelectedItem.Index = 2 Then
      Set frmAddEditPersonContacts.TempCollection = m_Employee.Contacts
      frmAddEditPersonContacts.ParentShowMode = ShowMode
      frmAddEditPersonContacts.ShowMode = SHOW_ADD
      frmAddEditPersonContacts.HeaderText = MapText("เพื่มสถานที่ติดต่อ")
      Load frmAddEditPersonContacts
      frmAddEditPersonContacts.Show 1

      OKClick = frmAddEditPersonContacts.OKClick

      Unload frmAddEditPersonContacts
      Set frmAddEditPersonContacts = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Employee.Contacts)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      Set frmAddEditPersonCards.TempCollection = m_Employee.Cards
      frmAddEditPersonCards.ParentShowMode = ShowMode
      frmAddEditPersonCards.ShowMode = SHOW_ADD
      frmAddEditPersonCards.HeaderText = MapText("เพื่มเอกสารราชการ")
      Load frmAddEditPersonCards
      frmAddEditPersonCards.Show 1

      OKClick = frmAddEditPersonCards.OKClick

      Unload frmAddEditPersonCards
      Set frmAddEditPersonCards = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Employee.Cards)
         GridEX1.Rebind
      End If
   
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
     Set frmAddEditPersonEmpWorked.TempCollection = m_Employee.EmpWorked
      frmAddEditPersonEmpWorked.ParentShowMode = ShowMode
      frmAddEditPersonEmpWorked.ShowMode = SHOW_ADD
      frmAddEditPersonEmpWorked.HeaderText = MapText("เพิ่มประวัติการทำงาน")
      Load frmAddEditPersonEmpWorked
      frmAddEditPersonEmpWorked.Show 1

      OKClick = frmAddEditPersonEmpWorked.OKClick

      Unload frmAddEditPersonEmpWorked
      Set frmAddEditPersonEmpWorked = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Employee.EmpWorked)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      Set frmAddEditPersonEmpEducation.TempCollection = m_Employee.EmpEducation
      frmAddEditPersonEmpEducation.ParentShowMode = ShowMode
      frmAddEditPersonEmpEducation.ShowMode = SHOW_ADD
      frmAddEditPersonEmpEducation.HeaderText = MapText("เพื่มประวัติการศึกษา")
      Load frmAddEditPersonEmpEducation
      frmAddEditPersonEmpEducation.Show 1

      OKClick = frmAddEditPersonEmpEducation.OKClick

      Unload frmAddEditPersonEmpEducation
      Set frmAddEditPersonEmpEducation = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Employee.EmpEducation)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
      Set frmAddEditPersonEmpDependency.TempCollection = m_Employee.EmpDependency
      frmAddEditPersonEmpDependency.ParentShowMode = ShowMode
      frmAddEditPersonEmpDependency.ShowMode = SHOW_ADD
      frmAddEditPersonEmpDependency.HeaderText = MapText("เพิ่มผู้เกี่ยวข้อง")
      Load frmAddEditPersonEmpDependency
      frmAddEditPersonEmpDependency.Show 1

      OKClick = frmAddEditPersonEmpDependency.OKClick

      Unload frmAddEditPersonEmpDependency
      Set frmAddEditPersonEmpDependency = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Employee.EmpDependency)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 7 Then
      Set frmAddEditPersonEmpChild.TempCollection = m_Employee.EmpChild
      frmAddEditPersonEmpChild.ParentShowMode = ShowMode
      frmAddEditPersonEmpChild.ShowMode = SHOW_ADD
      frmAddEditPersonEmpChild.HeaderText = MapText("เพื่มข้อมูลบุตร")
      Load frmAddEditPersonEmpChild
      frmAddEditPersonEmpChild.Show 1

      OKClick = frmAddEditPersonEmpChild.OKClick

      Unload frmAddEditPersonEmpChild
      Set frmAddEditPersonEmpChild = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Employee.EmpChild)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 8 Then
      Set frmAddEditPersonEmpHealty.TempCollection = m_Employee.EmpHealty
      frmAddEditPersonEmpHealty.ParentShowMode = ShowMode
      frmAddEditPersonEmpHealty.ShowMode = SHOW_ADD
      frmAddEditPersonEmpHealty.HeaderText = MapText("เพื่มประวัติการรักษาพยาบาล")
      Load frmAddEditPersonEmpHealty
      frmAddEditPersonEmpHealty.Show 1

      OKClick = frmAddEditPersonEmpHealty.OKClick

      Unload frmAddEditPersonEmpHealty
      Set frmAddEditPersonEmpHealty = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Employee.EmpHealty)
         GridEX1.Rebind
      End If
   
       
   End If

   If OKClick Then
      m_HasModify = True
   End If
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
   If TabStrip1.SelectedItem.Index = 2 Then
      If ID1 <= 0 Then
         m_Employee.Contacts.Remove (ID2)
      Else
         m_Employee.Contacts.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Employee.Contacts)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      If ID1 <= 0 Then
         m_Employee.Cards.Remove (ID2)
      Else
         m_Employee.Cards.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Employee.Cards)
      GridEX1.Rebind
      m_HasModify = True
   
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      If ID1 <= 0 Then
         m_Employee.EmpWorked.Remove (ID2)
      Else
         m_Employee.EmpWorked.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Employee.EmpWorked)
      GridEX1.Rebind
      m_HasModify = True
   
   
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      If ID1 <= 0 Then
         m_Employee.EmpEducation.Remove (ID2)
      Else
         m_Employee.EmpEducation.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Employee.EmpEducation)
      GridEX1.Rebind
      m_HasModify = True
   
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
      If ID1 <= 0 Then
         m_Employee.EmpDependency.Remove (ID2)
      Else
         m_Employee.EmpDependency.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Employee.EmpDependency)
      GridEX1.Rebind
      m_HasModify = True
   
   ElseIf TabStrip1.SelectedItem.Index = 7 Then
      If ID1 <= 0 Then
         m_Employee.EmpChild.Remove (ID2)
      Else
         m_Employee.EmpChild.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Employee.EmpChild)
      GridEX1.Rebind
      m_HasModify = True
   
   ElseIf TabStrip1.SelectedItem.Index = 8 Then
      If ID1 <= 0 Then
         m_Employee.EmpHealty.Remove (ID2)
      Else
         m_Employee.EmpHealty.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Employee.EmpHealty)
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
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   id = Val(GridEX1.Value(2))
   OKClick = False
   If TabStrip1.SelectedItem.Index = 2 Then
     Set frmAddEditPersonContacts.TempCollection = m_Employee.Contacts
      frmAddEditPersonContacts.id = id
      frmAddEditPersonContacts.ShowMode = SHOW_EDIT
      frmAddEditPersonContacts.HeaderText = MapText("แก้ไขสถานที่ติดต่อ")
      Load frmAddEditPersonContacts
      frmAddEditPersonContacts.Show 1

      OKClick = frmAddEditPersonContacts.OKClick

      Unload frmAddEditPersonContacts
      Set frmAddEditPersonContacts = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Employee.Contacts)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
     Set frmAddEditPersonCards.TempCollection = m_Employee.Cards
      frmAddEditPersonCards.id = id
      frmAddEditPersonCards.ShowMode = SHOW_EDIT
      frmAddEditPersonCards.HeaderText = MapText("แก้ไขสถานที่ติดต่อ")
      Load frmAddEditPersonCards
      frmAddEditPersonCards.Show 1

      OKClick = frmAddEditPersonCards.OKClick

      Unload frmAddEditPersonCards
      Set frmAddEditPersonCards = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Employee.Cards)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
     Set frmAddEditPersonEmpWorked.TempCollection = m_Employee.EmpWorked
      frmAddEditPersonEmpWorked.id = id
      frmAddEditPersonEmpWorked.ShowMode = SHOW_EDIT
      frmAddEditPersonEmpWorked.HeaderText = MapText("แก้ไขประวัติการทำงาน")
      Load frmAddEditPersonEmpWorked
      frmAddEditPersonEmpWorked.Show 1

      OKClick = frmAddEditPersonEmpWorked.OKClick

      Unload frmAddEditPersonEmpWorked
      Set frmAddEditPersonEmpWorked = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Employee.EmpWorked)
         GridEX1.Rebind
      End If
   
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
     Set frmAddEditPersonEmpEducation.TempCollection = m_Employee.EmpEducation
      frmAddEditPersonEmpEducation.id = id
      frmAddEditPersonEmpEducation.ShowMode = SHOW_EDIT
      frmAddEditPersonEmpEducation.HeaderText = MapText("แก้ไขประวัติการศึกษา")
      Load frmAddEditPersonEmpEducation
      frmAddEditPersonEmpEducation.Show 1

      OKClick = frmAddEditPersonEmpEducation.OKClick

      Unload frmAddEditPersonEmpEducation
      Set frmAddEditPersonEmpEducation = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Employee.EmpEducation)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
     Set frmAddEditPersonEmpDependency.TempCollection = m_Employee.EmpDependency
      frmAddEditPersonEmpDependency.id = id
      frmAddEditPersonEmpDependency.ShowMode = SHOW_EDIT
      frmAddEditPersonEmpDependency.HeaderText = MapText("แก้ไขผู้เกี่ยวข้อง")
      Load frmAddEditPersonEmpDependency
      frmAddEditPersonEmpDependency.Show 1

      OKClick = frmAddEditPersonEmpDependency.OKClick

      Unload frmAddEditPersonEmpDependency
      Set frmAddEditPersonEmpDependency = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Employee.EmpDependency)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 7 Then
     Set frmAddEditPersonEmpChild.TempCollection = m_Employee.EmpChild
      frmAddEditPersonEmpChild.id = id
      frmAddEditPersonEmpChild.ShowMode = SHOW_EDIT
      frmAddEditPersonEmpChild.HeaderText = MapText("แก้ไขข้อมูลบุตร")
      Load frmAddEditPersonEmpChild
      frmAddEditPersonEmpChild.Show 1

      OKClick = frmAddEditPersonEmpChild.OKClick

      Unload frmAddEditPersonEmpChild
      Set frmAddEditPersonEmpChild = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Employee.EmpChild)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 8 Then
     Set frmAddEditPersonEmpHealty.TempCollection = m_Employee.EmpHealty
      frmAddEditPersonEmpHealty.id = id
      frmAddEditPersonEmpHealty.ShowMode = SHOW_EDIT
      frmAddEditPersonEmpHealty.HeaderText = MapText("แก้ไขประวัติการรักษาพยาบาล")
      Load frmAddEditPersonEmpHealty
      frmAddEditPersonEmpHealty.Show 1

      OKClick = frmAddEditPersonEmpHealty.OKClick

      Unload frmAddEditPersonEmpHealty
      Set frmAddEditPersonEmpHealty = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Employee.EmpHealty)
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

Private Sub cmdPictureAdd_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Picture Files (*.jpg, *.gif)|*.jpg;*.gif"
   dlgAdd.DialogTitle = "Select Picture to Add to Database"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   m_HasModify = True
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadSex(cboSex)
      Call LoadWorkStatus(cboWorkStatus)
      Call LoadPosition(cboPosition)
      Call LoadBloodGroup(cboBloodGroup)
      Call LoadResignReason(cboCauseOut)
      Call LoadReligious(cboSadsana)
      Call LoadMarital(cboMary)
      Call LoadMilitary(cbotahan)
      Call LoadBankAccount(cboBank)
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Employee.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_Employee.QueryFlag = 0
         Call QueryData(False)
      End If
      
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
   
   Set m_Employee = Nothing
   Set m_Employees = Nothing
End Sub



Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   '''''''Debug.Print ColIndex & " " & NewColWidth
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
  GridEX1.Columns.Item(1).Visible = False
   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
  GridEX1.Columns.Item(2).Visible = False

   Set Col = GridEX1.Columns.add '3
   Col.Width = 1500
   Col.Caption = MapText("บ้านเลขที่")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2500
   Col.Caption = MapText("ซอย")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 1000
   Col.Caption = MapText("หมู่")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 2500
   Col.Caption = MapText("หมู่บ้าน")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 2500
   Col.Caption = MapText("ถนน")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 1980
   Col.Caption = MapText("ตำบล")
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1980
   Col.Caption = MapText("อำเภอ")
   Set Col = GridEX1.Columns.add '10
   Col.Width = 1980
   Col.Caption = MapText("จังหวัด")
   Set Col = GridEX1.Columns.add '11
   Col.Width = 1980
   Col.Caption = MapText("รหัสไปรษณีย์")
   Set Col = GridEX1.Columns.add '12
   Col.Width = 1980
   Col.Caption = MapText("ประเทศ")
End Sub

Private Sub InitGrid3()
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
GridEX1.Columns.Item(1).Visible = False
   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
GridEX1.Columns.Item(2).Visible = False
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2500
   Col.Caption = MapText("หมายเลขเอกสาร")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2500
   Col.Caption = MapText("วันที่ออกบัตร")

 Set Col = GridEX1.Columns.add '5
   Col.Width = 2500
   Col.Caption = MapText("วันที่หมดอายุ")
   Set Col = GridEX1.Columns.add '6
   Col.Width = 2500
   Col.Caption = MapText("สถานที่ออกบัตร")

   Set Col = GridEX1.Columns.add '7
   Col.Width = 2500
   Col.Caption = MapText("ประเภทบัตร")
End Sub

Private Sub InitGrid4()
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
GridEX1.Columns.Item(1).Visible = False
   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
GridEX1.Columns.Item(2).Visible = False
   Set Col = GridEX1.Columns.add '3
   Col.Width = 3500
   Col.Caption = MapText("สถานที่ทำงาน")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2500
   Col.Caption = MapText("จากวันที่")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2500
   Col.Caption = MapText("ถึงวันที่")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 3500
   'Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ตำแหน่ง")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 2500
   Col.Caption = MapText("สาเหตุที่ออก")

End Sub

Private Sub InitGrid5()
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
GridEX1.Columns.Item(1).Visible = False
   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   GridEX1.Columns.Item(2).Visible = False
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1500
   Col.Caption = MapText("วุฒิ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3500
   Col.Caption = MapText("สถาบัน")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2500
   Col.Caption = MapText("จากวันที่")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 2500
   Col.Caption = MapText("ถึงวันที่")

   Set Col = GridEX1.Columns.add '7
   Col.Width = 2500
   Col.Caption = MapText("วิชาหลัก")
   Set Col = GridEX1.Columns.add '8
   Col.Width = 1000
   Col.Caption = MapText("GPA")
   End Sub
Private Sub InitGrid6()
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
GridEX1.Columns.Item(1).Visible = False
   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
GridEX1.Columns.Item(2).Visible = False
   Set Col = GridEX1.Columns.add '3
   Col.Width = 3500
   Col.Caption = MapText("ชื่อ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3500
   Col.Caption = MapText("นามสกุล")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2500
   Col.Caption = MapText("ความสัมพันธ์")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 2500
   Col.Caption = MapText("วันเกิด")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 2500
   Col.Caption = MapText("หมายเลขติดต่อ")
   End Sub
Private Sub InitGrid7()
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
GridEX1.Columns.Item(1).Visible = False
   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
GridEX1.Columns.Item(2).Visible = False
   Set Col = GridEX1.Columns.add '3
   Col.Width = 3500
   Col.Caption = MapText("ชื่อ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3500
   Col.Caption = MapText("นามสกุล")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2500
   Col.Caption = MapText("ชื่อเล่น")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 2500
   Col.Caption = MapText("วันเกิด")
   Set Col = GridEX1.Columns.add '7
   Col.Width = 2500
   Col.Caption = MapText("หมายเลขติดต่อ")
   
   End Sub
Private Sub InitGrid8()
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
GridEX1.Columns.Item(1).Visible = False
   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
GridEX1.Columns.Item(2).Visible = False
   Set Col = GridEX1.Columns.add '3
   Col.Width = 3500
   Col.Caption = MapText("โรคที่รักษา")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3500
   Col.Caption = MapText("โรงพยาบาล")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2500
   Col.Caption = MapText("จากวันที่")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 2500
   Col.Caption = MapText("ถึงวันที่")
   End Sub


Private Sub GetTotalPrice()
'Dim Ii As CTransferItem
'Dim Sum As Double
'
'   Sum = 0
'   For Each Ii In m_Employee.TransferItems
'      If Ii.Flag <> "D" Then
'         Sum = Sum + CDbl(Format(Ii.ExportItem.EXPORT_AVG_PRICE, "0.00")) * CDbl(Format(Ii.ExportItem.EXPORT_AMOUNT, "0.00"))
'      End If
'   Next Ii
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   fraPerson.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   'fraPerson.Visible = True
   Call InitNormalLabel(lblPersonName, MapText("ชื่อพนักงาน"))
   Call InitNormalLabel(lblPersonLname, MapText("นามสกุล"))
   Call InitNormalLabel(lblPersonSex, MapText("เพศ"))
   Call InitNormalLabel(lblPersonCode, MapText("รหัสพนักงาน"))
   
   Call InitNormalLabel(lblBirthDate, MapText("วันเกิด"))
   Call InitNormalLabel(lblWorkStatus, MapText("สถานะการทำงาน"))
   Call InitNormalLabel(lblInDate, MapText("วันที่เข้างาน"))
   Call InitNormalLabel(lblPassDate, MapText("วันที่ผ่านงาน"))
   
   Call InitNormalLabel(lblWeight, MapText("น้ำหนัก"))
   Call InitNormalLabel(lblWeightUnit, MapText("กิโลกรัม"))
   Call InitNormalLabel(lblOutDate, MapText("วันที่ออก"))
   Call InitNormalLabel(lblHight, MapText("ส่วนสูง"))
   
      Call InitNormalLabel(lblHightUnit, MapText("เซนติเมตร"))
   Call InitNormalLabel(lblCauseOut, MapText("สาเหตุที่ออก"))
   Call InitNormalLabel(lblBloodGroup, MapText("กรุ๊ปเลือด"))
   Call InitNormalLabel(lblNickName, MapText("ชื่อเล่น"))

   Call InitNormalLabel(lblNation, MapText("สัญชาติ"))
   Call InitNormalLabel(lblSalary, MapText("เงินเดือน"))
   Call InitNormalLabel(lblSalaryUnit, MapText("บาท"))
   Call InitNormalLabel(lblOt, MapText("โอที"))

   Call InitNormalLabel(lblOtUnit, MapText("บาท/ชม."))
   Call InitNormalLabel(lblInterNational, MapText("เชื้อชาติ"))
   Call InitNormalLabel(lblPosition, MapText("ตำแหน่ง"))
   Call InitNormalLabel(lblSadsana, MapText("ศาสนา"))

   Call InitNormalLabel(lblMary, MapText("สถานะสมรส"))
   Call InitNormalLabel(lbltahan, MapText("สถานะการทหาร"))
   Call InitNormalLabel(lblBank, MapText("ธนาคาร"))
   Call InitNormalLabel(lblBranchBank, MapText("สาขา"))

Call InitNormalLabel(lblDepositNo, MapText("เลขที่บัญชี"))
   Call InitNormalLabel(lblRok, MapText("โรคประจำตัว"))
   Call InitNormalLabel(lblHomeNo, MapText("เบอร์บ้าน"))
   Call InitNormalLabel(lblMobileNo, MapText("เบอร์มือถือ"))
   Call InitNormalLabel(lblEmail, MapText("อีเมล์"))
   
   
   Call txtPersonCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
 Call txtPersonName.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
Call txtPersonLname.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)

Call txtWeight.SetTextLenType(TEXT_STRING, glbSetting.PORT_TYPE)
 Call txtHight.SetTextLenType(TEXT_STRING, glbSetting.PORT_TYPE)
Call txtNickName.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)

Call txtNation.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
 Call txtSalary.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.NAME_TYPE)
Call txtOt.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.NAME_TYPE)

Call txtInterNational.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
 Call txtDepositNo.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
Call txtBranchBank.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)

Call txtRok.SetTextLenType(TEXT_STRING, glbSetting.ADDRESS_TYPE)
 Call txtHomeNo.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
Call txtMobileNo.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
Call txtEmail.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   
   Call InitCheckBox(chkPersonOut, MapText("พนักงานนอก"))
   chkPersonOut.Visible = False
   Call InitCombo(cboSex)
   Call InitCombo(cboWorkStatus)
Call InitCombo(cboCauseOut)
Call InitCombo(cboBloodGroup)
Call InitCombo(cboPosition)
Call InitCombo(cboSadsana)
   Call InitCombo(cboMary)
Call InitCombo(cbotahan)
Call InitCombo(cboBank)

   
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
  ' TabStrip1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("ข้อมูลทั่วไป")
   TabStrip1.Tabs.add().Caption = MapText("สถานที่ติดต่อ")
   TabStrip1.Tabs.add().Caption = MapText("ข้อมูลเอกสารราชการ")
   TabStrip1.Tabs.add().Caption = MapText("ประวัติการทำงาน")
   TabStrip1.Tabs.add().Caption = MapText("ประวัติการศึกษา")
   TabStrip1.Tabs.add().Caption = MapText("ผู้เกี่ยวข้อง")
   TabStrip1.Tabs.add().Caption = MapText("ข้อมูลบุตร")
   TabStrip1.Tabs.add().Caption = MapText(" ประวัติการรักษาพยาบาล ")
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
   Set m_Employee = New CEmployee
   Set m_Employees = New Collection
   Set TempCollection = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   If TabStrip1.SelectedItem.Index = 2 Then
      RowBuffer.RowStyle = RowBuffer.Value(11)
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      RowBuffer.RowStyle = RowBuffer.Value(7)
ElseIf TabStrip1.SelectedItem.Index = 4 Then
      RowBuffer.RowStyle = RowBuffer.Value(7)
ElseIf TabStrip1.SelectedItem.Index = 5 Then
      RowBuffer.RowStyle = RowBuffer.Value(7)
   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 2 Then
      
      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CEmpAddress
      Dim Addr As CAddress
           Set CR = GetItem(m_Employee.Contacts, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
      Set Addr = CR.Addresses(1)
      Values(1) = Addr.ADDRESS_ID
      Values(2) = RealIndex
      Values(3) = Addr.HOME
      Values(4) = Addr.SOI
      Values(5) = Addr.MOO
      Values(6) = Addr.VILLAGE
      Values(7) = Addr.ROAD
      Values(8) = Addr.DISTRICT
      Values(9) = Addr.AMPHUR
      Values(10) = Addr.PROVINCE
      Values(11) = Addr.ZIPCODE
      Values(12) = Addr.COUNTRY_NAME
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
     If m_Employee.Cards Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CP As CEmployeeProof
      Dim Doc As CDocument
      If m_Employee.Cards.Count <= 0 Then
         Exit Sub
      End If
      Set CP = GetItem(m_Employee.Cards, RowIndex, RealIndex)
      If CP Is Nothing Then
         Exit Sub
      End If
      Set Doc = CP.Doc

      Values(1) = Doc.DOCUMENT_ID
      Values(2) = RealIndex
      Values(3) = Doc.DOCUMENT_NO
      Values(4) = DateToStringExt(Doc.ISSUE_DATE)
      Values(5) = DateToStringExt(Doc.EXPIRE_DATE)
      Values(6) = Doc.Address.AMPHUR
      Values(7) = Doc.DOCTYPE_NAME
  
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   If m_Employee.EmpWorked Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim EW As CEmpWorked
      If m_Employee.EmpWorked.Count <= 0 Then
         Exit Sub
      End If
      Set EW = GetItem(m_Employee.EmpWorked, RowIndex, RealIndex)
      If EW Is Nothing Then
         Exit Sub
      End If
'Dim ReasonName As EW.RESIGN_REASON
      Values(1) = EW.EMP_WORKED_ID
      Values(2) = RealIndex
      Values(3) = EW.WORK_PLACE
      Values(4) = DateToStringExt(EW.FROM_DATE)
      Values(5) = DateToStringExt(EW.TO_DATE)
      Values(6) = EW.EMP_POSITION
      Values(7) = EW.RESIGN_REASON_NAME

   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      If m_Employee.EmpEducation Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim EE As CEmpEducation
      If m_Employee.EmpEducation.Count <= 0 Then
         Exit Sub
      End If
      Set EE = GetItem(m_Employee.EmpEducation, RowIndex, RealIndex)
      If EE Is Nothing Then
         Exit Sub
      End If

      Values(1) = EE.EMP_EDUCATION_ID
      Values(2) = RealIndex
      Values(3) = EE.QUALIFICATION_NAME
      Values(4) = EE.INSTITUTE
      Values(5) = DateToStringExt(EE.FROM_DATE)
      Values(6) = DateToStringExt(EE.TO_DATE)
      Values(7) = EE.MASTER
      Values(8) = EE.SCORE
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
      If m_Employee.EmpDependency Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Ed As CEmpDependency
      If m_Employee.EmpDependency.Count <= 0 Then
         Exit Sub
      End If
      Set Ed = GetItem(m_Employee.EmpDependency, RowIndex, RealIndex)
      If Ed Is Nothing Then
         Exit Sub
      End If

      Values(1) = Ed.EMP_DEPENDENCY_ID
      Values(2) = RealIndex
      Values(3) = Ed.NAME.LONG_NAME
      Values(4) = Ed.NAME.LAST_NAME
      Values(5) = Ed.DEPENDENCY_NAME
      Values(6) = DateToStringExt(Ed.BIRTH_DATE)
      Values(7) = Ed.PHONE
   ElseIf TabStrip1.SelectedItem.Index = 7 Then
      If m_Employee.EmpChild Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim EC As CEmpChild
      If m_Employee.EmpChild.Count <= 0 Then
         Exit Sub
      End If
      Set EC = GetItem(m_Employee.EmpChild, RowIndex, RealIndex)
      If EC Is Nothing Then
         Exit Sub
      End If

      Values(1) = EC.EMP_CHILD_ID
      Values(2) = RealIndex
      Values(3) = EC.NAME.LONG_NAME
      Values(4) = EC.NAME.LAST_NAME
      Values(5) = EC.NAME.NICK_NAME
      Values(6) = DateToStringExt(EC.BIRTH_DATE)
      Values(7) = EC.PHONE
   ElseIf TabStrip1.SelectedItem.Index = 8 Then
      If m_Employee.EmpHealty Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim EH As CEmpHealty
      If m_Employee.EmpHealty.Count <= 0 Then
         Exit Sub
      End If
      Set EH = GetItem(m_Employee.EmpHealty, RowIndex, RealIndex)
      If EH Is Nothing Then
         Exit Sub
      End If

      Values(1) = EH.EMP_HEALTY_ID
      Values(2) = RealIndex
      Values(3) = EH.HEALT_DESC
      Values(4) = EH.HOSPITAL_NAME
      Values(5) = DateToStringExt(EH.FROM_DATE)
      Values(6) = DateToStringExt(EH.TO_DATE)
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
cmdAdd.Visible = True
cmdEdit.Visible = True
cmdDelete.Visible = True
   fraPerson.Visible = False
   fraGrid.Visible = False
   If TabStrip1.SelectedItem.Index = 1 Then
      fraPerson.Visible = True
      Call AllVisibleFalse
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   fraGrid.Visible = True
      Call InitGrid2
            GridEX1.ItemCount = CountItem(m_Employee.Contacts)
      GridEX1.Rebind
         ElseIf TabStrip1.SelectedItem.Index = 3 Then
   fraGrid.Visible = True
       Call InitGrid3
            GridEX1.ItemCount = CountItem(m_Employee.Cards)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   fraGrid.Visible = True
       Call InitGrid4
            GridEX1.ItemCount = CountItem(m_Employee.EmpWorked)
      GridEX1.Rebind
   
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   fraGrid.Visible = True
       Call InitGrid5
   GridEX1.ItemCount = CountItem(m_Employee.EmpEducation)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 6 Then
fraGrid.Visible = True
       Call InitGrid6
   GridEX1.ItemCount = CountItem(m_Employee.EmpDependency)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 7 Then
   fraGrid.Visible = True
       Call InitGrid7
       GridEX1.ItemCount = CountItem(m_Employee.EmpChild)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 8 Then
   fraGrid.Visible = True
       Call InitGrid8
       GridEX1.ItemCount = CountItem(m_Employee.EmpHealty)
      GridEX1.Rebind
   End If
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

Private Sub txtSender_Change()
   m_HasModify = True
End Sub

Private Sub txtTotal_Change()
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

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlEmployeeLookup_Change()
   m_HasModify = True
End Sub

Private Sub cboBusinessGroup_Click()
   m_HasModify = True
End Sub

Private Sub cboBusinessType_Click()
   m_HasModify = True
End Sub

Private Sub cboEnterpriseType_Click()
   m_HasModify = True
End Sub

Private Sub cboReason_Click()
   m_HasModify = True
End Sub

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub txtBranchBank_Change()
m_HasModify = True
End Sub

Private Sub txtDepositNo_Change()
m_HasModify = True
End Sub

Private Sub txtEmail_Change()
m_HasModify = True
End Sub

Private Sub txtHight_Change()
m_HasModify = True
End Sub

Private Sub txtHomeNo_Change()
m_HasModify = True
End Sub

Private Sub txtInterNational_Change()
m_HasModify = True
End Sub

Private Sub txtMobileNo_Change()
m_HasModify = True
End Sub

Private Sub txtNation_Change()
m_HasModify = True
End Sub

Private Sub txtNickName_Change()
m_HasModify = True
End Sub

Private Sub txtOt_Change()
m_HasModify = True
End Sub

Private Sub txtPersonCode_Change()
m_HasModify = True
End Sub

Private Sub txtPersonLname_Change()
m_HasModify = True
End Sub

Private Sub txtPersonName_Change()
m_HasModify = True
End Sub

Private Sub txtRok_Change()
m_HasModify = True
End Sub

Private Sub txtSalary_Change()
m_HasModify = True
End Sub

Private Sub txtWeight_Change()
m_HasModify = True
End Sub

Private Sub uctlBirthDate_HasChange()
m_HasModify = True
End Sub

Private Sub uctlInDate_HasChange()
m_HasModify = True
End Sub

Private Sub uctlOutDate_HasChange()
m_HasModify = True
End Sub

Private Sub uctlPassDate_HasChange()
m_HasModify = True
End Sub
