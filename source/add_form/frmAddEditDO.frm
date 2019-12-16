VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddEditDO 
   ClientHeight    =   8430
   ClientLeft      =   2580
   ClientTop       =   -105
   ClientWidth     =   11895
   Icon            =   "frmAddEditDO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000009&
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   1575
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   1635
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   10200
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   17992
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSFrame SSFrame3 
         Height          =   2565
         Left            =   120
         TabIndex        =   67
         Top             =   9120
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   4524
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.ComboBox cboHold4 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   1740
            Width           =   2355
         End
         Begin VB.ComboBox cboHold3 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   1290
            Width           =   2355
         End
         Begin VB.ComboBox cboHold2 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   840
            Width           =   2355
         End
         Begin VB.ComboBox cboHold1 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   390
            Width           =   2355
         End
         Begin prjFarmManagement.uctlTextBox txtHold1Desc 
            Height          =   435
            Left            =   7080
            TabIndex        =   70
            Top             =   390
            Width           =   4335
            _extentx        =   5001
            _extenty        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtHold1Amount 
            Height          =   435
            Left            =   4380
            TabIndex        =   69
            Top             =   390
            Width           =   1515
            _extentx        =   5001
            _extenty        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtHold2Desc 
            Height          =   435
            Left            =   7080
            TabIndex        =   73
            Top             =   840
            Width           =   4335
            _extentx        =   5001
            _extenty        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtHold2Amount 
            Height          =   435
            Left            =   4380
            TabIndex        =   72
            Top             =   840
            Width           =   1515
            _extentx        =   5001
            _extenty        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtHold3Desc 
            Height          =   435
            Left            =   7080
            TabIndex        =   76
            Top             =   1290
            Width           =   4335
            _extentx        =   5001
            _extenty        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtHold3Amount 
            Height          =   435
            Left            =   4380
            TabIndex        =   75
            Top             =   1290
            Width           =   1515
            _extentx        =   5001
            _extenty        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtHold4Desc 
            Height          =   435
            Left            =   7080
            TabIndex        =   79
            Top             =   1740
            Width           =   4335
            _extentx        =   5001
            _extenty        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtHold4Amount 
            Height          =   435
            Left            =   4380
            TabIndex        =   78
            Top             =   1740
            Width           =   1515
            _extentx        =   5001
            _extenty        =   767
         End
         Begin VB.Label lblHold4 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   90
            TabIndex        =   92
            Top             =   1740
            Width           =   1065
         End
         Begin VB.Label lblHold4Desc 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6030
            TabIndex        =   91
            Top             =   1770
            Width           =   945
         End
         Begin VB.Label lblHold4Amount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3570
            TabIndex        =   90
            Top             =   1770
            Width           =   705
         End
         Begin VB.Label lblHold3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   90
            TabIndex        =   89
            Top             =   1290
            Width           =   1065
         End
         Begin VB.Label lblHold3Desc 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6030
            TabIndex        =   88
            Top             =   1320
            Width           =   945
         End
         Begin VB.Label lblHold3Amount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3570
            TabIndex        =   87
            Top             =   1320
            Width           =   705
         End
         Begin VB.Label lblHold2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   90
            TabIndex        =   86
            Top             =   840
            Width           =   1065
         End
         Begin VB.Label lblHold2Desc 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6030
            TabIndex        =   85
            Top             =   870
            Width           =   945
         End
         Begin VB.Label lblHold2Amount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3570
            TabIndex        =   84
            Top             =   870
            Width           =   705
         End
         Begin VB.Label lblHold1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   83
            Top             =   390
            Width           =   1065
         End
         Begin VB.Label lblHold1Desc 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6030
            TabIndex        =   82
            Top             =   420
            Width           =   945
         End
         Begin VB.Label lblHold1Amount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3570
            TabIndex        =   81
            Top             =   420
            Width           =   705
         End
         Begin VB.Label Label12 
            Height          =   315
            Left            =   3360
            TabIndex        =   80
            Top             =   300
            Visible         =   0   'False
            Width           =   405
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   2565
         Left            =   5760
         TabIndex        =   44
         Top             =   9120
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   4524
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin prjFarmManagement.uctlTime uctlTime1 
            Height          =   405
            Left            =   7680
            TabIndex        =   55
            Top             =   2010
            Width           =   1125
            _extentx        =   1984
            _extenty        =   714
         End
         Begin prjFarmManagement.uctlDate uctlDueDate 
            Height          =   405
            Left            =   7680
            TabIndex        =   47
            Top             =   150
            Width           =   3855
            _extentx        =   6800
            _extenty        =   714
         End
         Begin prjFarmManagement.uctlTextBox txtPayment 
            Height          =   435
            Left            =   1740
            TabIndex        =   50
            Top             =   1050
            Width           =   4305
            _extentx        =   7594
            _extenty        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtNote 
            Height          =   435
            Left            =   1740
            TabIndex        =   48
            Top             =   600
            Width           =   1515
            _extentx        =   5001
            _extenty        =   767
         End
         Begin prjFarmManagement.uctlDate uctlShipDate 
            Height          =   405
            Left            =   7680
            TabIndex        =   51
            Top             =   1050
            Width           =   3855
            _extentx        =   6800
            _extenty        =   714
         End
         Begin prjFarmManagement.uctlTextBox txtPONo 
            Height          =   435
            Left            =   8640
            TabIndex        =   49
            Top             =   600
            Width           =   2805
            _extentx        =   4948
            _extenty        =   767
         End
         Begin prjFarmManagement.uctlTextLookup uctlSellByLookup 
            Height          =   435
            Left            =   1740
            TabIndex        =   52
            Top             =   1530
            Width           =   4305
            _extentx        =   9499
            _extenty        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtCredit 
            Height          =   435
            Left            =   1740
            TabIndex        =   46
            Top             =   150
            Width           =   1515
            _extentx        =   5001
            _extenty        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtTempDONo 
            Height          =   435
            Left            =   1740
            TabIndex        =   54
            Top             =   1980
            Width           =   2625
            _extentx        =   4630
            _extenty        =   767
         End
         Begin prjFarmManagement.uctlTime uctlTime2 
            Height          =   405
            Left            =   9060
            TabIndex        =   56
            Top             =   2010
            Width           =   1125
            _extentx        =   1984
            _extenty        =   714
         End
         Begin prjFarmManagement.uctlTextBox txtGeneration 
            Height          =   435
            Left            =   4680
            TabIndex        =   117
            Top             =   120
            Width           =   1515
            _extentx        =   5001
            _extenty        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtReference 
            Height          =   435
            Left            =   4680
            TabIndex        =   119
            Top             =   600
            Width           =   2475
            _extentx        =   4366
            _extenty        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtFarmName 
            Height          =   435
            Left            =   7680
            TabIndex        =   121
            Top             =   1560
            Width           =   3195
            _extentx        =   5636
            _extenty        =   767
         End
         Begin VB.Label lblFarmName 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6240
            TabIndex        =   120
            Top             =   1560
            Width           =   1395
         End
         Begin VB.Label lblReference 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3360
            TabIndex        =   118
            Top             =   600
            Width           =   1155
         End
         Begin VB.Label lblGeneration 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3360
            TabIndex        =   116
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Caption         =   "-"
            Height          =   315
            Left            =   8880
            TabIndex        =   114
            Top             =   2070
            Width           =   105
         End
         Begin VB.Label lblInOutTime 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5910
            TabIndex        =   113
            Top             =   2070
            Width           =   1605
         End
         Begin VB.Label lblTempDONo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   60
            TabIndex        =   94
            Top             =   2010
            Width           =   1605
         End
         Begin VB.Label Label11 
            Height          =   315
            Left            =   3360
            TabIndex        =   66
            Top             =   240
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label lblCredit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            TabIndex        =   65
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label lblSellBy 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   30
            TabIndex        =   60
            Top             =   1590
            Width           =   1635
         End
         Begin VB.Label lblPoNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7200
            TabIndex        =   59
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label lblDeliveryPlace 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   210
            TabIndex        =   58
            Top             =   1140
            Width           =   1395
         End
         Begin VB.Label lblShipment 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5640
            TabIndex        =   57
            Top             =   1110
            Width           =   1995
         End
         Begin VB.Label lblNote 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            TabIndex        =   53
            Top             =   630
            Width           =   1395
         End
         Begin VB.Label lblDueDate 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6120
            TabIndex        =   45
            Top             =   240
            Width           =   1395
         End
      End
      Begin VB.ComboBox cboPackageType 
         Height          =   315
         Left            =   9690
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2670
         Width           =   1755
      End
      Begin VB.ComboBox cboEnpAddress 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2220
         Width           =   9585
      End
      Begin VB.ComboBox cboCustomerAddress 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1770
         Width           =   9585
      End
      Begin VB.ComboBox cboAccount 
         Height          =   315
         Left            =   9090
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   2355
      End
      Begin prjFarmManagement.uctlTextLookup uctlCustomerLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   1320
         Width           =   5385
         _extentx        =   9499
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6570
         TabIndex        =   1
         Top             =   870
         Width           =   3855
         _extentx        =   6800
         _extenty        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   13
         Top             =   4680
         Width           =   11595
         _ExtentX        =   20452
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
      Begin prjFarmManagement.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   870
         Width           =   2190
         _extentx        =   5001
         _extenty        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   120
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2205
         Left            =   150
         TabIndex        =   14
         Top             =   5160
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   3889
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
         Column(1)       =   "frmAddEditDO.frx":27A2
         Column(2)       =   "frmAddEditDO.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditDO.frx":290E
         FormatStyle(2)  =   "frmAddEditDO.frx":2A6A
         FormatStyle(3)  =   "frmAddEditDO.frx":2B1A
         FormatStyle(4)  =   "frmAddEditDO.frx":2BCE
         FormatStyle(5)  =   "frmAddEditDO.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditDO.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtTotalAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   9
         Top             =   2640
         Width           =   1605
         _extentx        =   2831
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNetTotal 
         Height          =   435
         Left            =   6030
         TabIndex        =   10
         Top             =   2640
         Width           =   1575
         _extentx        =   2778
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDiscount 
         Height          =   435
         Left            =   1860
         TabIndex        =   11
         Top             =   3090
         Width           =   1605
         _extentx        =   2831
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtIncludeDiscount 
         Height          =   435
         Left            =   6030
         TabIndex        =   33
         Top             =   3090
         Width           =   1575
         _extentx        =   2778
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCashDiscount 
         Height          =   435
         Left            =   1860
         TabIndex        =   38
         Top             =   4590
         Visible         =   0   'False
         Width           =   1605
         _extentx        =   2831
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLeft 
         Height          =   435
         Left            =   6030
         TabIndex        =   12
         Top             =   3540
         Width           =   1575
         _extentx        =   2778
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCashDiscountAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   62
         Top             =   3540
         Width           =   1605
         _extentx        =   2831
         _extenty        =   767
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   2595
         Left            =   -240
         TabIndex        =   95
         Top             =   9840
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   4577
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.ComboBox cboPaymentType 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   96
            Top             =   270
            Width           =   2325
         End
         Begin VB.ComboBox cboBank 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   99
            Top             =   720
            Width           =   4035
         End
         Begin VB.ComboBox cboBankBranch 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Top             =   1170
            Width           =   4035
         End
         Begin prjFarmManagement.uctlTextBox txtCheckNo 
            Height          =   435
            Left            =   7350
            TabIndex        =   97
            Top             =   270
            Width           =   3465
            _extentx        =   5001
            _extenty        =   767
         End
         Begin prjFarmManagement.uctlDate uctlCheckDate 
            Height          =   405
            Left            =   7350
            TabIndex        =   100
            Top             =   720
            Width           =   3855
            _extentx        =   6800
            _extenty        =   714
         End
         Begin VB.Label lblBankBranch 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   60
            TabIndex        =   98
            Top             =   1230
            Width           =   1395
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   60
            TabIndex        =   105
            Top             =   750
            Width           =   1395
         End
         Begin VB.Label lblPaymentType 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   104
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label lblCheckDate 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5820
            TabIndex        =   103
            Top             =   780
            Width           =   1395
         End
         Begin VB.Label lblCheckNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5580
            TabIndex        =   101
            Top             =   330
            Width           =   1665
         End
      End
      Begin prjFarmManagement.uctlTextBox txtDipRcp 
         Height          =   435
         Left            =   6030
         TabIndex        =   106
         Top             =   3960
         Width           =   1575
         _extentx        =   2778
         _extenty        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalRcp 
         Height          =   435
         Left            =   1860
         TabIndex        =   107
         Top             =   3960
         Width           =   1605
         _extentx        =   2831
         _extenty        =   767
      End
      Begin VB.Label lblMsg 
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   150
         TabIndex        =   124
         Top             =   7440
         Width           =   11475
      End
      Begin Threed.SSCheck chkPostFlag 
         Height          =   435
         Left            =   9480
         TabIndex        =   123
         Top             =   3960
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "chkPostFlag"
      End
      Begin Threed.SSCommand cmdEditDeliveyCost 
         Height          =   525
         Left            =   5160
         TabIndex        =   122
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   4080
         TabIndex        =   112
         Top             =   870
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDO.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label Label15 
         Height          =   315
         Left            =   7740
         TabIndex        =   111
         Top             =   4020
         Width           =   585
      End
      Begin VB.Label lblDipRcp 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4290
         TabIndex        =   110
         Top             =   4050
         Width           =   1635
      End
      Begin VB.Label lblTotalRcp 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   109
         Top             =   4080
         Width           =   1425
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   3510
         TabIndex        =   108
         Top             =   4020
         Width           =   765
      End
      Begin VB.Label lblPackageType 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8460
         TabIndex        =   93
         Top             =   2700
         Width           =   1125
      End
      Begin VB.Label Label9 
         Height          =   315
         Left            =   3570
         TabIndex        =   64
         Top             =   3600
         Width           =   765
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   300
         TabIndex        =   63
         Top             =   3630
         Width           =   1425
      End
      Begin VB.Label Label5 
         Height          =   315
         Left            =   3540
         TabIndex        =   61
         Top             =   4650
         Visible         =   0   'False
         Width           =   765
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   6810
         TabIndex        =   18
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdCustomer 
         Height          =   405
         Left            =   7260
         TabIndex        =   4
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDO.frx":3250
         ButtonStyle     =   3
      End
      Begin VB.Label lblCashDiscount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   42
         Top             =   4710
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label10 
         Height          =   315
         Left            =   3540
         TabIndex        =   41
         Top             =   3660
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblLeft 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4710
         TabIndex        =   40
         Top             =   3630
         Width           =   1275
      End
      Begin VB.Label Label8 
         Height          =   315
         Left            =   7680
         TabIndex        =   39
         Top             =   3600
         Width           =   585
      End
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   37
         Top             =   3210
         Width           =   1695
      End
      Begin VB.Label Label6 
         Height          =   315
         Left            =   3540
         TabIndex        =   36
         Top             =   3180
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblIncludeDiscount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4710
         TabIndex        =   35
         Top             =   3180
         Width           =   1275
      End
      Begin VB.Label Label3 
         Height          =   315
         Left            =   7680
         TabIndex        =   34
         Top             =   3150
         Width           =   585
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   10500
         TabIndex        =   2
         Top             =   870
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "chkCommit"
      End
      Begin VB.Label lblEnpAddress 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   2310
         Width           =   1635
      End
      Begin VB.Label lblCustomerAddress 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   1860
         Width           =   1635
      End
      Begin VB.Label lblAccountNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7890
         TabIndex        =   30
         Top             =   1410
         Width           =   1095
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   150
         TabIndex        =   29
         Top             =   1380
         Width           =   1635
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   7710
         TabIndex        =   28
         Top             =   2700
         Width           =   585
      End
      Begin VB.Label lblNetTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5010
         TabIndex        =   27
         Top             =   2730
         Width           =   915
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   3540
         TabIndex        =   26
         Top             =   2730
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5040
         TabIndex        =   25
         Top             =   930
         Width           =   1395
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   19
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDO.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   20
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDO.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   17
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDO.frx":3B9E
         ButtonStyle     =   3
      End
      Begin VB.Label lblTotalAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   23
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   22
         Top             =   930
         Width           =   1665
      End
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   0
      TabIndex        =   115
      Top             =   0
      Width           =   1395
   End
End
Attribute VB_Name = "frmAddEditDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean

Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_BillingDoc As CBillingDoc
Private m_Customers As Collection
Private m_Employees As Collection
Private m_Resources As Collection
Private m_CustomerPictures As Collection
Private m_Cd As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public Area As Long
Public DocumentType As Long
Public ReceiptType As Long

Private Programowner As String
Private FileName As String
Private m_SumUnit As Double
Private AllowSave As Boolean
Private EditDeliveryCostFlag As Boolean
Private DocAdd As Long

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_BillingDoc.BILLING_DOC_ID = ID
      If Not glbDaily.QueryBillingDoc(m_BillingDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_BillingDoc.PopulateFromRS(1, m_Rs)
      
      uctlDocumentDate.ShowDate = m_BillingDoc.DOCUMENT_DATE
      txtDocumentNo.Text = m_BillingDoc.DOCUMENT_NO
      If Area = 1 Then
         uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_BillingDoc.CUSTOMER_ID)
         cboAccount.ListIndex = IDToListIndex(cboAccount, m_BillingDoc.ACCOUNT_ID)
      ElseIf Area = 2 Then
         uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, m_BillingDoc.SUPPLIER_ID)
         cboAccount.ListIndex = -1
      End If
      cboCustomerAddress.ListIndex = IDToListIndex(cboCustomerAddress, m_BillingDoc.BILLING_ADDRESS_ID)
      cboEnpAddress.ListIndex = IDToListIndex(cboEnpAddress, m_BillingDoc.ENTERPRISE_ADDRESS_ID)
      txtTotalAmount.Text = Format(m_BillingDoc.TOTAL_AMOUNT, "0.00")
      txtCashDiscount.Text = Format(m_BillingDoc.CD_PERCENT, "0.00")
      txtCashDiscountAmount.Text = Format(m_BillingDoc.CD_AMOUNT, "0.00")
      txtTotalRcp.Text = Format(m_BillingDoc.TOTAL_RCP, "0.00")
      uctlSellByLookup.MyCombo.ListIndex = IDToListIndex(uctlSellByLookup.MyCombo, m_BillingDoc.ACCEPT_BY)
      txtCredit.Text = m_BillingDoc.Credit
      uctlDueDate.ShowDate = m_BillingDoc.DUE_DATE
      cboPackageType.ListIndex = IDToListIndex(cboPackageType, m_BillingDoc.PACKAGE_TYPE)
      txtTempDONo.Text = m_BillingDoc.TEMP_DO_NO
      txtGeneration.Text = m_BillingDoc.GENERATION
      txtReference.Text = m_BillingDoc.REFERENCE
      txtFarmName.Text = m_BillingDoc.FARM_NAME
      
      uctlShipDate.ShowDate = m_BillingDoc.SHIPMENT
      txtPayment.Text = m_BillingDoc.PAYMENT_DESC
      txtNote.Text = m_BillingDoc.NOTE
      txtPONo.Text = m_BillingDoc.REF
      chkCommit.Value = FlagToCheck(m_BillingDoc.OLD_COMMIT_FLAG)
      chkCommit.Enabled = (m_BillingDoc.OLD_COMMIT_FLAG = "N")
      chkPostFlag.Value = FlagToCheck(m_BillingDoc.POST_FLAG)
      
     If DocumentType = 1 Then
     If Not VerifyAccessRight("LEDGER_SELL" & "_" & DocumentType & "_" & "POST", "กำหนดเอกสารสมบูรณ์", 2) Then
        chkPostFlag.Enabled = False
      Else
        chkPostFlag.Enabled = (m_BillingDoc.POST_FLAG = "N")
     End If
   End If
      Call EnableDisableButton(True)
      
      txtCheckNo.Text = m_BillingDoc.CHECK_NO
      uctlCheckDate.ShowDate = m_BillingDoc.CHECK_DATE
      cboPaymentType.ListIndex = IDToListIndex(cboPaymentType, m_BillingDoc.PAYMENT_TYPE)
      cboBank.ListIndex = IDToListIndex(cboBank, m_BillingDoc.BANK_ID)
      cboBankBranch.ListIndex = IDToListIndex(cboBankBranch, m_BillingDoc.BBRANCH_ID)
      
      uctlTime1.HR = HOUR(m_BillingDoc.ENTRY_DATE)
      uctlTime1.MI = Minute(m_BillingDoc.ENTRY_DATE)
      uctlTime2.HR = HOUR(m_BillingDoc.EXIT_DATE)
      uctlTime2.MI = Minute(m_BillingDoc.EXIT_DATE)
      
      Call LoadDoPartItem
      Call ShowBulkHole
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

Private Sub PopulateGuiID(BD As CBillingDoc)
Dim Di As CDoItem

   For Each Di In BD.DoItems
      If Di.Flag = "A" Then
         Di.LINK_ID = GetNextGuiID(BD)
      End If
   Next Di
End Sub

Private Function GetNextGuiID(BD As CBillingDoc) As Long
Dim Di As CDoItem
Dim MaxId As Long

   MaxId = 0
   For Each Di In BD.DoItems
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   GetNextGuiID = MaxId + 1
End Function

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc
Dim Pm As CPayment
Dim TempCreditLimit As Double
Dim CheckMaxCredit As Boolean
Dim TempUserName As String
   
   If ShowMode = SHOW_EDIT Then
      If Area = 1 Then
         If Not VerifyAccessRight("LEDGER_SELL" & "_" & DocumentType & "_" & "EDIT", "แก้ไข") Then
            frmVerifyAccRight.AccName = "LEDGER_SELL" & "_" & DocumentType & "_" & "EDIT"
            frmVerifyAccRight.AccDesc = "แก้ไข"
            Load frmVerifyAccRight
            frmVerifyAccRight.Show 1
            
            If frmVerifyAccRight.GrantRight Then
               Unload frmVerifyAccRight
               Set frmVerifyAccRight = Nothing
            Else
               Unload frmVerifyAccRight
               Set frmVerifyAccRight = Nothing
               Call EnableForm(Me, True)
               Exit Function
            End If
         End If
      End If
   End If
   
   If Not VerifyTextControl(lblDocumentNo, txtDocumentNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblDocumentDate, uctlDocumentDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblTotalAmount, txtTotalAmount, True) Then
      Exit Function
   End If
   If Not VerifyDate(lblDueDate, uctlDueDate, True) Then
      Exit Function
   End If

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   'ตรวจสอบการแก้ไขเอกสารที่ออกบิลแล้ว
   If m_BillingDoc.POST_FLAG = "Y" Then
         glbErrorLog.LocalErrorMsg = MapText("เอกสารใบนี้ได้ทำการออกบิลขายเรียบร้อยแล้ว หากต้องการเปลี่ยนแปลงเอกสารต้องให้ผู้ควบคุม อนุมัติก่อน")
         glbErrorLog.ShowUserError
      
         frmVerifyAccRight.AccName = "LEDGER_SELL_" & DocumentType & "_" & "MANAGE"
         frmVerifyAccRight.AccDesc = "สามารถอนุมัติการเปลี่ยนแปลงเอกสารได้"
         Load frmVerifyAccRight
         frmVerifyAccRight.Show 1
         If frmVerifyAccRight.GrantRight Then
            TempUserName = frmVerifyAccRight.UserName
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
            m_BillingDoc.APPROVE_MANAGE_NAME = TempUserName
            chkPostFlag.Enabled = True
         Else
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
            Exit Function
         End If
   End If

   If Not CheckUniqueNs(DO_PLAN_UNIQUE, txtDocumentNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      txtDocumentNo.Text = ""
      DocAdd = DocAdd + 1
      Call cmdAuto_Click
      Exit Function
   End If
   
   If DocumentType = 1 Then
      Dim Cs As CCustomer
      Dim Cus_ID As Long
      Dim SumDB As Double
      Dim TotalSale As Double
      Cus_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
      Set Cs = m_Customers(Trim(str(Cus_ID)))
      
      If DocumentType = 1 Then
         If Not (AllowSave) Then
            glbErrorLog.LocalErrorMsg = MapText("ยังไม่มีข้อมูล ใบปะหน้าบัญชี และเวลาได้เกินกำหนดแล้วไม่สามารถสร้างเอกสารใบขายเชื่อได้ " & vbCrLf & "  กรุณาเพิ่มข้อมูลใบปะหน้าบัญชีเพิ่มให้สามารถใช้ Funtion นี้ได้")
            glbErrorLog.ShowUserError
            Exit Function
         End If
      End If
   
   'เพิ่มวงเงินราย Week 01/11/2560 โดยพี่มณ
   If Cs.WEEK_CREDIT_LIMIT > 0 Then
      TotalSale = CheckWeekCreditLimit(uctlDocumentDate.ShowDate)
      TotalSale = TotalSale + (Val(txtNetTotal.Text) - Val(txtDiscount.Text) - Val(txtCashDiscountAmount.Text))
      If TotalSale > Cs.WEEK_CREDIT_LIMIT Then
         glbErrorLog.LocalErrorMsg = "ยอดหนี้มากกว่าวงเงินรายสัปดาห์ (จำนวน " & FormatNumber(TotalSale - Cs.WEEK_CREDIT_LIMIT) & " ) "
         glbErrorLog.ShowUserError
         Exit Function
      End If
   End If
   'เพิ่มวงเงินราย Week 01/11/2560 โดยพี่มณ

 'เพิ่มการรับเงินสดจากลูกค้า 15/02/2561 โดยพี่สุชาย

   If Cs.CASH_FLAG = "Y" Then
      glbErrorLog.LocalErrorMsg = "ลูกค้ารายนี้ชำระด้วยเงินสด" & vbNewLine & "กรุณาให้เจ้าหน้าที่บัญชีมารับเงินและอนุมัติเอกสารใบนี้"
      glbErrorLog.ShowUserError
      frmVerifyAccRight.AccName = "CREDIT_CASH-PAY"
      frmVerifyAccRight.AccDesc = "สามารถอนุมัติการซื้อขายด้วยเงินสด"
      Load frmVerifyAccRight
      frmVerifyAccRight.Show 1
          If frmVerifyAccRight.GrantRight Then
            TempUserName = frmVerifyAccRight.UserName
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
            m_BillingDoc.CASH_FLAG = "Y"
            m_BillingDoc.APPROVE_NAME = TempUserName
         Else
            Unload frmVerifyAccRight
            Set frmVerifyAccRight = Nothing
            Exit Function
         End If

Else
'If Not EditDeliveryCostFlag Then 'ถ้าไม่ใช่การแก้ไขค่าขนส่ง
      'จิวเอาออกก่อน วันที่ 15/09/2558
         SumDB = 0
         If Cs.CHECK_CREDIT_FLAG = "Y" Then
            CheckMaxCredit = True
            If Cs.SUSPEND_SALES = "Y" Or Cs.CREDIT_LIMIT > 0 Or Cs.MAX_CREDIT > 0 Then
               If Cs.SUSPEND_SALES = "N" Then
                  If Cs.CREDIT_LIMIT > 0 Then   'เช็ควงเงิน
                     TempCreditLimit = CheckCreditLimit
                     SumDB = (Cs.CREDIT_LIMIT - TempCreditLimit) - (Val(txtNetTotal.Text) - Val(txtDiscount.Text) - Val(txtCashDiscountAmount.Text))
                  End If
                  
                  'เช็ควงวัน
                  CheckMaxCredit = CheckMaxCreditDueDate(Cs)
               End If
               If Cs.SUSPEND_SALES = "Y" Or SumDB < 0 Or (Not CheckMaxCredit) Then
                  If Cs.SUSPEND_SALES = "Y" Then
                     glbErrorLog.LocalErrorMsg = "ระงับการขายชั่วคราว"
                  ElseIf SumDB < 0 And (Not CheckMaxCredit) Then
                     glbErrorLog.LocalErrorMsg = "ยอดหนี้มากกว่าวงเงิน (จำนวน " & (-SumDB) & " ) และมียอดเกินวงวันสูงสุด"
                  ElseIf SumDB < 0 Then
                     glbErrorLog.LocalErrorMsg = "ยอดหนี้มากกว่าวงเงิน (จำนวน " & (-SumDB) & " ) "
                  Else
                     glbErrorLog.LocalErrorMsg = "มียอดเกินวงวันสูงสุด"
                  End If
                  glbErrorLog.ShowUserError
      
                     frmVerifyAccRight.AccName = "CREDIT_CONTROL"
                     frmVerifyAccRight.AccDesc = "สามารถเปลี่ยนแปลงเครดิตซื้อขาย"
                     Load frmVerifyAccRight
                     frmVerifyAccRight.Show 1
         
                     If frmVerifyAccRight.GrantRight Then
                        TempUserName = frmVerifyAccRight.UserName
                        Unload frmVerifyAccRight
                        Set frmVerifyAccRight = Nothing
                        m_BillingDoc.APPROVE_NAME = TempUserName
                        
                     Else
                        Unload frmVerifyAccRight
                        Set frmVerifyAccRight = Nothing
                        Exit Function
                     End If
      
   
               End If
      
               If Cs.SUSPEND_SALES = "N" Then
                  m_BillingDoc.OLD_CREDIT_AMOUNT = TempCreditLimit
               Else
                  m_BillingDoc.OLD_CREDIT_AMOUNT = CheckCreditLimit
               End If
            End If
         End If
'      End If 'end If Not EditDeliveryCostFlag Then
'      EditDeliveryCostFlag = False
   End If
  End If
   
   If CountItem(m_BillingDoc.Payments) <= 0 And DocumentType = 2 Then
      glbErrorLog.LocalErrorMsg = "กรุณาใส่การชำระเงินใหถูกต้องครบถ้วน"
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   m_BillingDoc.AddEditMode = ShowMode
   m_BillingDoc.BILLING_DOC_ID = ID
   m_BillingDoc.DOCUMENT_DATE = uctlDocumentDate.ShowDate
   m_BillingDoc.DOCUMENT_NO = txtDocumentNo.Text
   If Area = 1 Then
      m_BillingDoc.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
      m_BillingDoc.ACCOUNT_ID = cboAccount.ItemData(Minus2Zero(cboAccount.ListIndex))
   ElseIf Area = 2 Then
      m_BillingDoc.SUPPLIER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
      m_BillingDoc.ACCOUNT_ID = -1
   End If
   m_BillingDoc.DOCUMENT_TYPE = DocumentType
   m_BillingDoc.RECEIPT_TYPE = ReceiptType
   m_BillingDoc.BILLING_ADDRESS_ID = cboCustomerAddress.ItemData(Minus2Zero(cboCustomerAddress.ListIndex))
   m_BillingDoc.ENTERPRISE_ADDRESS_ID = cboEnpAddress.ItemData(Minus2Zero(cboEnpAddress.ListIndex))
   m_BillingDoc.EXCEPTION_FLAG = "N"
   m_BillingDoc.ACCEPT_BY = uctlSellByLookup.MyCombo.ItemData(Minus2Zero(uctlSellByLookup.MyCombo.ListIndex))
   m_BillingDoc.COMMIT_FLAG = Check2Flag(chkCommit.Value)
   m_BillingDoc.POST_FLAG = Check2Flag(chkPostFlag.Value)
   m_BillingDoc.DISCOUNT_AMOUNT = Val(txtDiscount.Text)
   m_BillingDoc.CD_PERCENT = Val(txtCashDiscount.Text)
   m_BillingDoc.CD_AMOUNT = Val(txtCashDiscountAmount.Text)
   m_BillingDoc.TOTAL_AMOUNT = Val(txtTotalAmount.Text)
   m_BillingDoc.TOTAL_PRICE = Val(txtNetTotal.Text)
   m_BillingDoc.Credit = Val(txtCredit.Text)
   m_BillingDoc.TOTAL_RCP = Val(txtTotalRcp.Text)
   m_BillingDoc.DUE_DATE = uctlDueDate.ShowDate
   m_BillingDoc.REF = txtPONo.Text
   m_BillingDoc.PACKAGE_TYPE = cboPackageType.ItemData(Minus2Zero(cboPackageType.ListIndex))
   m_BillingDoc.TEMP_DO_NO = txtTempDONo.Text
   m_BillingDoc.GENERATION = txtGeneration.Text
   m_BillingDoc.REFERENCE = txtReference.Text
   m_BillingDoc.FARM_NAME = txtFarmName.Text
   
   m_BillingDoc.CHECK_NO = txtCheckNo.Text
   m_BillingDoc.CHECK_DATE = uctlCheckDate.ShowDate
   m_BillingDoc.PAYMENT_TYPE = cboPaymentType.ItemData(Minus2Zero(cboPaymentType.ListIndex))
   m_BillingDoc.BANK_ID = cboBank.ItemData(Minus2Zero(cboBank.ListIndex))
   If cboBankBranch.ListIndex > 0 Then
      m_BillingDoc.BBRANCH_ID = cboBankBranch.ItemData(Minus2Zero(cboBankBranch.ListIndex))
   End If
      
   m_BillingDoc.ENTRY_DATE = uctlDocumentDate.ShowDate
   m_BillingDoc.ENTRY_DATE = DateAdd("h", uctlTime1.HR, m_BillingDoc.ENTRY_DATE)
   m_BillingDoc.ENTRY_DATE = DateAdd("n", uctlTime1.MI, m_BillingDoc.ENTRY_DATE)
   m_BillingDoc.EXIT_DATE = uctlDocumentDate.ShowDate
   m_BillingDoc.EXIT_DATE = DateAdd("h", uctlTime2.HR, m_BillingDoc.EXIT_DATE)
   m_BillingDoc.EXIT_DATE = DateAdd("n", uctlTime2.MI, m_BillingDoc.EXIT_DATE)
            
  m_BillingDoc.SHIPMENT = uctlShipDate.ShowDate
   m_BillingDoc.NOTE = txtNote.Text
   m_BillingDoc.PAYMENT_DESC = txtPayment.Text
   
   Call PopulateGuiID(m_BillingDoc)
   Call glbDaily.GenerateExtraDiscount(m_BillingDoc)
   
   Call EnableForm(Me, False)
   
   If DocumentType = 1 Then   ' ใบส่งของขาย
      Call glbDaily.DO2InventoryDoc(m_BillingDoc, Ivd, Area, 10)
   ElseIf DocumentType = 2 Then  'ใบเสร็จขาย
      Call glbDaily.DO2InventoryDoc(m_BillingDoc, Ivd, Area, 21)
      Call glbDaily.DO2Payment(m_BillingDoc, Pm)
   ElseIf DocumentType = 7 Then   ' ใบส่งของซื้อ
      Call glbDaily.DO2InventoryDoc(m_BillingDoc, Ivd, Area, 100)
   End If
   
   If (m_BillingDoc.COMMIT_FLAG = "Y") Then
      If m_BillingDoc.OLD_COMMIT_FLAG <> "Y" Then
         Call glbDaily.TriggerCommit(Ivd.ImportExports)
         If Not glbDaily.VerifyStockBalance(Ivd.ImportExports, glbErrorLog) Then
            Call EnableForm(Me, True)
            Exit Function
         End If
         
      End If
   End If
   
   Call glbDaily.StartTransaction
   Call EditStatusFlagInInventoryWHDoc(m_BillingDoc)
   Call EditStatusFlagInBillingDoc(m_BillingDoc)
   
   If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         SaveData = False
         Call glbDaily.RollbackTransaction
         Call EnableForm(Me, True)
         Exit Function
   End If
   
   If DocumentType = 2 Then
      If Not glbDaily.AddEditPayment(Pm, IsOK, False, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         SaveData = False
         Call glbDaily.RollbackTransaction
         Call EnableForm(Me, True)
         Exit Function
      End If
      m_BillingDoc.PAYMENT_ID = Pm.PAYMENT_ID
   End If
   
   Call PopulateBulkHole
   
   m_BillingDoc.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
   If Not glbDaily.AddEditBillingDoc(m_BillingDoc, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
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

Private Sub ShowBulkHole()
On Error Resume Next
Dim Bh As CBulkHole
   
   Set Bh = m_BillingDoc.BulkHoles(1)
   cboHold1.ListIndex = IDToListIndex(cboHold1, Bh.PART_ITEM_ID)
   txtHold1Amount.Text = Bh.HOLE_AMOUNT
   txtHold1Desc.Text = Bh.NOTE

   Set Bh = m_BillingDoc.BulkHoles(2)
   cboHold2.ListIndex = IDToListIndex(cboHold2, Bh.PART_ITEM_ID)
   txtHold2Amount.Text = Bh.HOLE_AMOUNT
   txtHold2Desc.Text = Bh.NOTE

   Set Bh = m_BillingDoc.BulkHoles(3)
   cboHold3.ListIndex = IDToListIndex(cboHold3, Bh.PART_ITEM_ID)
   txtHold3Amount.Text = Bh.HOLE_AMOUNT
   txtHold3Desc.Text = Bh.NOTE

   Set Bh = m_BillingDoc.BulkHoles(4)
   cboHold4.ListIndex = IDToListIndex(cboHold3, Bh.PART_ITEM_ID)
   txtHold4Amount.Text = Bh.HOLE_AMOUNT
   txtHold4Desc.Text = Bh.NOTE
End Sub

Private Sub PopulateBulkHole()
Dim Bh As CBulkHole

   For Each Bh In m_BillingDoc.BulkHoles
      Bh.Flag = "D"
   Next Bh
   
   Set Bh = New CBulkHole
   Bh.Flag = "A"
   Bh.PART_ITEM_ID = cboHold1.ItemData(Minus2Zero(cboHold1.ListIndex))
   Bh.HOLE_AMOUNT = Val(txtHold1Amount.Text)
   Bh.NOTE = txtHold1Desc.Text
   Call m_BillingDoc.BulkHoles.add(Bh)
   Set Bh = Nothing
   
   Set Bh = New CBulkHole
   Bh.Flag = "A"
   Bh.PART_ITEM_ID = cboHold2.ItemData(Minus2Zero(cboHold2.ListIndex))
   Bh.HOLE_AMOUNT = Val(txtHold2Amount.Text)
   Bh.NOTE = txtHold2Desc.Text
   Call m_BillingDoc.BulkHoles.add(Bh)
   Set Bh = Nothing
   
   Set Bh = New CBulkHole
   Bh.Flag = "A"
   Bh.PART_ITEM_ID = cboHold3.ItemData(Minus2Zero(cboHold3.ListIndex))
   Bh.HOLE_AMOUNT = Val(txtHold3Amount.Text)
   Bh.NOTE = txtHold3Desc.Text
   Call m_BillingDoc.BulkHoles.add(Bh)
   Set Bh = Nothing
   
   Set Bh = New CBulkHole
   Bh.Flag = "A"
   Bh.PART_ITEM_ID = cboHold4.ItemData(Minus2Zero(cboHold4.ListIndex))
   Bh.HOLE_AMOUNT = Val(txtHold4Amount.Text)
   Bh.NOTE = txtHold4Desc.Text
   Call m_BillingDoc.BulkHoles.add(Bh)
   Set Bh = Nothing
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

Private Sub cboAccount_Click()
   m_HasModify = True
End Sub

Private Sub cboAccount_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboBank_Click()
Dim BankID As Long

   BankID = cboBank.ItemData(Minus2Zero(cboBank.ListIndex))
   If BankID > 0 Then
      Call LoadBankBranch(cboBankBranch, , BankID)
   End If

   m_HasModify = True
End Sub

Private Sub cboBankBranch_Click()
   m_HasModify = True
End Sub

Private Sub cboCustomerAddress_Click()
   m_HasModify = True
End Sub

Private Sub cboCustomerAddress_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboEnpAddress_Click()
   m_HasModify = True
End Sub

Private Sub cboEnpAddress_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboHold1_Click()
   m_HasModify = True
End Sub

Private Sub cboHold2_Click()
   m_HasModify = True
End Sub

Private Sub cboHold3_Click()
   m_HasModify = True
End Sub

Private Sub cboHold4_Click()
   m_HasModify = True
End Sub

Private Sub cboPackageType_Click()
   m_HasModify = True
End Sub

Private Sub cboPackageType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboPaymentType_Click()
   m_HasModify = True
End Sub

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Public Sub RefreshGrid()
   Call GetTotalPrice

   GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
   GridEX1.Rebind
End Sub

Private Sub chkPostFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   If Area = 1 Then
      If Not VerifyCombo(lblAccountNo, cboAccount) Then
         Exit Sub
      End If
      If Not VerifyDate(lblDocumentDate, uctlDocumentDate) Then
         Exit Sub
      End If
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      If Area = 1 Then
         Set oMenu = New cPopupMenu
         lMenuChosen = oMenu.AddMenu(glbGuiConfigs.DOAddMenuItems)
         Set oMenu = Nothing
         If lMenuChosen = 0 Then
            Exit Sub
         End If
      Else
         lMenuChosen = 1
      End If

       If lMenuChosen = 1 Then
         If Area = 1 Then
            frmAddEditDoItem.AccountID = cboAccount.ItemData(cboAccount.ListIndex)
         End If

         frmAddEditDoItem.DocumentDate = uctlDocumentDate.ShowDate
         frmAddEditDoItem.SubscriberID = -1
         frmAddEditDoItem.Area = Area
         frmAddEditDoItem.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
         Set frmAddEditDoItem.TempCollection = m_BillingDoc.DoItems
         frmAddEditDoItem.ParentShowMode = ShowMode
         frmAddEditDoItem.ShowMode = SHOW_ADD
         frmAddEditDoItem.HeaderText = MapText("เพิ่มรายการใบส่งสินค้า")
    
         Load frmAddEditDoItem
         frmAddEditDoItem.Show 1
   
         OKClick = frmAddEditDoItem.OKClick
   
         Unload frmAddEditDoItem
         Set frmAddEditDoItem = Nothing
   
         If OKClick Then
            Call GetTotalPrice
   
            GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
            GridEX1.Rebind
         End If
      ElseIf lMenuChosen = 2 Then
         If Area = 1 Then
            frmAddEditDoItemEx.AccountID = cboAccount.ItemData(cboAccount.ListIndex)
         Else
            glbErrorLog.LocalErrorMsg = "ฟังก์ชันนี้ไม่สนับสนุนในส่วนงานซื้อ"
            glbErrorLog.ShowUserError
            Exit Sub
         End If
         Set frmAddEditDoItemEx.ParentForm = Me
         frmAddEditDoItemEx.SubscriberID = -1
         frmAddEditDoItemEx.Area = Area
         frmAddEditDoItemEx.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
         Set frmAddEditDoItemEx.TempCollection = m_BillingDoc.DoItems
         frmAddEditDoItemEx.ParentShowMode = ShowMode
         frmAddEditDoItemEx.ShowMode = SHOW_ADD
         frmAddEditDoItemEx.HeaderText = MapText("เพิ่มรายการใบส่งสินค้า")
         Load frmAddEditDoItemEx
         frmAddEditDoItemEx.Show 1

         OKClick = frmAddEditDoItemEx.OKClick

         Unload frmAddEditDoItemEx
         Set frmAddEditDoItemEx = Nothing

         If OKClick Then
            Call GetTotalPrice
   
            GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
            GridEX1.Rebind
         End If
      ElseIf lMenuChosen = 4 Then
         frmAddPOItem.AccountID = cboAccount.ItemData(Minus2Zero(cboAccount.ListIndex))
         Set frmAddPOItem.TempCollection = m_BillingDoc.DoItems
         frmAddPOItem.ShowMode = SHOW_ADD
         frmAddPOItem.HeaderText = MapText("เพิ่มรายการใบส่งของจากใบ PO")
         
         Load frmAddPOItem
         frmAddPOItem.Show 1
   
         OKClick = frmAddPOItem.OKClick
   
         Unload frmAddPOItem
         Set frmAddPOItem = Nothing
   
         If OKClick Then
            Call GetTotalPrice
   
            GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
            GridEX1.Rebind
         End If
      ElseIf lMenuChosen = 5 Then
         frmAddQuoatationItem.AccountID = cboAccount.ItemData(Minus2Zero(cboAccount.ListIndex))
         Set frmAddQuoatationItem.TempCollection = m_BillingDoc.DoItems
         frmAddQuoatationItem.ShowMode = SHOW_ADD
         frmAddQuoatationItem.HeaderText = MapText("เพิ่มรายการใบส่งของจากใบเสนอราคา")
         
         Load frmAddQuoatationItem
         frmAddQuoatationItem.Show 1
   
         OKClick = frmAddQuoatationItem.OKClick
   
         Unload frmAddQuoatationItem
         Set frmAddQuoatationItem = Nothing
   
         If OKClick Then
            Call GetTotalPrice
   
            GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
            GridEX1.Rebind
         End If
      ElseIf lMenuChosen = 6 Then
         Set frmAddSOItem.TempCollection = m_BillingDoc.DoItems
         frmAddSOItem.Area = 2 'มาจาก ใบส่งของ
         frmAddSOItem.DocumentDate = uctlDocumentDate.ShowDate
         frmAddSOItem.CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
         frmAddSOItem.ShowMode = SHOW_ADD
         frmAddSOItem.HeaderText = MapText("เพิ่มรายการใบส่งของจากใบ SO")
        
         
         Load frmAddSOItem
         frmAddSOItem.Show 1
   
         OKClick = frmAddSOItem.OKClick
   
         Unload frmAddSOItem
         Set frmAddSOItem = Nothing
   
         If OKClick Then
          Call GetTotalPrice
         GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
         GridEX1.Rebind
         End If
        ElseIf lMenuChosen = 7 Then
            Set frmAddSOItem.T_CBillingDoc = m_BillingDoc
            Set frmAddSOItem.TempCollection = m_BillingDoc.DoItems
            frmAddSOItem.Area = 4
            frmAddSOItem.DOCUMENT_TYPE = 2000
            frmAddSOItem.DocumentDate = uctlDocumentDate.ShowDate
            frmAddSOItem.CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
            frmAddSOItem.CustomerCode = uctlCustomerLookup.MyTextBox.Text
            frmAddSOItem.ShowMode = SHOW_ADD
            frmAddSOItem.HeaderText = MapText("เพิ่มรายการใบส่งของจากใบขึ้นอาหาร BAG")
   
            Load frmAddSOItem
            frmAddSOItem.Show 1
      
            OKClick = frmAddSOItem.OKClick
      
            Unload frmAddSOItem
            Set frmAddSOItem = Nothing
      
            If OKClick Then
             Call GetTotalPrice
            GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
            GridEX1.Rebind
            End If
            
           txtNote.Text = m_BillingDoc.NOTE
           txtPONo.Text = m_BillingDoc.REF
           txtReference.Text = m_BillingDoc.REFERENCE
           txtPayment.Text = m_BillingDoc.PAYMENT_DESC
           txtTempDONo.Text = m_BillingDoc.TEMP_DO_NO
      ElseIf lMenuChosen = 9 Then
            Set frmAddSOItem.T_CBillingDoc = m_BillingDoc
            Set frmAddSOItem.TempCollection = m_BillingDoc.DoItems
            frmAddSOItem.Area = 5
            frmAddSOItem.DOCUMENT_TYPE = 2001
            frmAddSOItem.DocumentDate = uctlDocumentDate.ShowDate
            frmAddSOItem.CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
            frmAddSOItem.CustomerCode = uctlCustomerLookup.MyTextBox.Text
            frmAddSOItem.ShowMode = SHOW_ADD
            frmAddSOItem.HeaderText = MapText("เพิ่มรายการใบส่งของจากใบขึ้นอาหาร BULK")
   
            Load frmAddSOItem
            frmAddSOItem.Show 1
      
            OKClick = frmAddSOItem.OKClick
      
            Unload frmAddSOItem
            Set frmAddSOItem = Nothing
      
            If OKClick Then
             Call GetTotalPrice
            GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
            GridEX1.Rebind
            End If
            
           txtNote.Text = m_BillingDoc.NOTE
           txtPONo.Text = m_BillingDoc.REF
           txtReference.Text = m_BillingDoc.REFERENCE
           txtPayment.Text = m_BillingDoc.PAYMENT_DESC
           txtTempDONo.Text = m_BillingDoc.TEMP_DO_NO
         ElseIf lMenuChosen = 11 Then
'            Set frmAddSOItem.T_CBillingDoc = m_BillingDoc
'            Set frmAddSOItem.TempCollection = m_BillingDoc.DoItems
'            frmAddSOItem.Area = 5
'            frmAddSOItem.DOCUMENT_TYPE = 2001
'            frmAddSOItem.DocumentDate = uctlDocumentDate.ShowDate
'            frmAddSOItem.CustomerId = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
'            frmAddSOItem.CustomerCode = uctlCustomerLookup.MyTextBox.Text
'            frmAddSOItem.ShowMode = SHOW_ADD
'            frmAddSOItem.HeaderText = MapText("เพิ่มรายการใบส่งของจากใบขึ้นอาหาร BULK")
'
'            Load frmAddSOItem
'            frmAddSOItem.Show 1
'
'            OKClick = frmAddSOItem.OKClick
'
'            Unload frmAddSOItem
'            Set frmAddSOItem = Nothing
'
'            If OKClick Then
'             Call GetTotalPrice
'            GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
'            GridEX1.Rebind
'            End If
'
'           txtNote.Text = m_BillingDoc.NOTE
'           txtPONo.Text = m_BillingDoc.REF
'           txtReference.Text = m_BillingDoc.REFERENCE
'           txtPayment.Text = m_BillingDoc.PAYMENT_DESC
'           txtTempDONo.Text = m_BillingDoc.TEMP_DO_NO
     ElseIf lMenuChosen = 13 Then
            Set frmAddRQItem.T_CBillingDoc = m_BillingDoc
            Set frmAddRQItem.TempCollection = m_BillingDoc.DoItems
            frmAddRQItem.DOCUMENT_TYPE = 3
            frmAddRQItem.DocumentDate = uctlDocumentDate.ShowDate
            frmAddRQItem.CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
            frmAddRQItem.ShowMode = SHOW_ADD
            frmAddRQItem.HeaderText = MapText("เพิ่มรายการใบส่งฝากขาย")
   
            Load frmAddRQItem
            frmAddRQItem.Show 1
      
            OKClick = frmAddRQItem.OKClick
      
            Unload frmAddRQItem
            Set frmAddRQItem = Nothing
      
            If OKClick Then
             Call GetTotalPrice
            GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
            GridEX1.Rebind
            End If
            
           txtNote.Text = m_BillingDoc.NOTE
           txtPONo.Text = m_BillingDoc.REF
           txtReference.Text = m_BillingDoc.REFERENCE
           txtPayment.Text = m_BillingDoc.PAYMENT_DESC
           txtTempDONo.Text = m_BillingDoc.TEMP_DO_NO
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If Area = 2 Then
         glbErrorLog.LocalErrorMsg = "ฟังก์ชันนี้ไม่สนับสนุนในส่วนงานซื้อ"
         glbErrorLog.ShowUserError
         Exit Sub
      End If
      frmAddEditBillingDiscount.Area = Area
      frmAddEditBillingDiscount.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
      Set frmAddEditBillingDiscount.TempCollection = m_BillingDoc.BillingDiscounts
      Set frmAddEditBillingDiscount.TempCollection2 = m_BillingDoc.DoItems
      frmAddEditBillingDiscount.ParentShowMode = ShowMode
      frmAddEditBillingDiscount.ShowMode = SHOW_ADD
      frmAddEditBillingDiscount.HeaderText = MapText("เพิ่มรายการส่วนลด")
      Load frmAddEditBillingDiscount
      frmAddEditBillingDiscount.Show 1

      OKClick = frmAddEditBillingDiscount.OKClick

      Unload frmAddEditBillingDiscount
      Set frmAddEditBillingDiscount = Nothing

      If OKClick Then
         Call GetTotalPrice

         GridEX1.ItemCount = CountItem(m_BillingDoc.BillingDiscounts)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      frmAddEditCashTran.Area = Area
      Set frmAddEditCashTran.ParentForm = Me
      frmAddEditCashTran.HeaderText = "เพิ่มรายการการชำระเงิน"
      frmAddEditCashTran.ShowMode = SHOW_ADD
      Set frmAddEditCashTran.TempCollection = m_BillingDoc.Payments
      Load frmAddEditCashTran
      frmAddEditCashTran.Show 1
      
      OKClick = frmAddEditCashTran.OKClick
      
      Unload frmAddEditCashTran
      Set frmAddEditCashTran = Nothing
   
      If OKClick Then
         m_HasModify = True
         
         GridEX1.ItemCount = CountItem(m_BillingDoc.Payments)
         Call GridEX1.Rebind
         
         Call GetTotalPrice
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
'Dim No As String
'
'   If Trim(txtDocumentNo.Text) = "" Then
'      Call glbDatabaseMngr.GenerateNumber(DO_NUMBER, No, glbErrorLog)
'      txtDocumentNo.Text = No
'   End If
End Sub
Private Function GetDocumentNo(DocNoType As Long) As String
Dim No As String
Dim DOC_ID As Long
Dim Cd As CConfigDoc
Dim TempStr As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim I As Long
Dim ServerDateTime As String

   DOC_ID = IV_DO
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
'         GetDocumentNo = GetDocumentNo & TempStr & Cd.GetFieldValue("CODE2")
'         TempStr = ""
         If Cd.GetFieldValue("MONTH_TYPE") = 1 Then
            TempStr2 = Format(Month(Now), "00")
         End If
'         GetDocumentNo = GetDocumentNo & TempStr2 & Cd.GetFieldValue("CODE3")
'         TempStr = ""
         For I = 1 To Cd.GetFieldValue("DIGIT_AMOUNT")
            TempStr3 = TempStr3 & "0"
         Next I

         If Cd.GetFieldValue("AUTO_BEGIN_FLAG") = "Y" Then
               If CheckNewMounth And CheckUniqueNs(DO_PLAN_UNIQUE, TempStr2 & "-" & Format(1, TempStr3) & "-" & TempStr, ID) Then
'                  GetDocumentNo = GetDocumentNo & Format(1, TempStr3) 'เริ่มจาก 1 เสมอ
                  TempStr3 = Format(1, TempStr3) 'เริ่มจาก 1 เสมอ
                  m_BillingDoc.RUNNING_NO = 1
               Else
'                  GetDocumentNo = GetDocumentNo & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr3)
                  GetDocumentNo = TempStr2 & "-" & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr3) & "-" & TempStr
                 m_BillingDoc.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
               End If
          Else
'               GetDocumentNo = GetDocumentNo & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr3)
                GetDocumentNo = TempStr2 & "-" & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr3) & "-" & TempStr
                m_BillingDoc.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
          End If
          m_BillingDoc.CONFIG_DOC_TYPE = DOC_ID
      Else
         GetDocumentNo = ""
      End If
      
   If Not CheckUniqueNs(DO_PLAN_UNIQUE, GetDocumentNo, ID) Then
      txtDocumentNo.Text = ""
      DocAdd = DocAdd + 1
      GetDocumentNo = GetDocumentNo(DocumentType)
   End If
      
   End If
End Function

Private Sub cmdCustomer_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim OKClick As Boolean
Dim TempCol As Collection
Dim Cs As CCustomer

   Set TempCol = New Collection
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("ค้นหา", "-", "เพิ่มข้อมูลใหม่", "-", "ตรวจสอบ Credit")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   If lMenuChosen = 1 Then
      Set frmQueryCustomer.TempCollection = TempCol
      frmQueryCustomer.ShowMode = SHOW_ADD
      Load frmQueryCustomer
      frmQueryCustomer.Show 1
      
      OKClick = frmQueryCustomer.OKClick
      
      Unload frmQueryCustomer
      Set frmQueryCustomer = Nothing
      
      If OKClick Then
         Set Cs = TempCol(1)
         uctlCustomerLookup.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookup.MyCombo, Cs.CUSTOMER_ID)
         m_HasModify = True
      End If
   ElseIf lMenuChosen = 3 Then
      frmAddEditCustomer.ShowMode = SHOW_ADD
      frmAddEditCustomer.HeaderText = MapText("เพิ่มลูกค้า")
      Load frmAddEditCustomer
      frmAddEditCustomer.Show 1
      
      OKClick = frmAddEditCustomer.OKClick
      Call EnableForm(Me, False)
      If Area = 1 Then
         Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
         Set uctlCustomerLookup.MyCollection = m_Customers
      ElseIf Area = 2 Then
         Call LoadSupplier(uctlCustomerLookup.MyCombo, m_Customers)
         Set uctlCustomerLookup.MyCollection = m_Customers
      End If
      Call EnableForm(Me, True)
      
      Unload frmAddEditCustomer
      Set frmAddEditCustomer = Nothing
   ElseIf lMenuChosen = 5 Then
      Screen.MousePointer = 11
      Call CheckCredit
      Screen.MousePointer = vbArrow
   End If
   
   Set TempCol = Nothing
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
         m_BillingDoc.DoItems.Remove (ID2)
      Else
         m_BillingDoc.DoItems.Item(ID2).Flag = "D"
      End If

      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If ID1 <= 0 Then
         m_BillingDoc.BillingDiscounts.Remove (ID2)
      Else
         m_BillingDoc.BillingDiscounts.Item(ID2).Flag = "D"
      End If

      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_BillingDoc.BillingDiscounts)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      If ID1 <= 0 Then
         m_BillingDoc.Payments.Remove (ID2)
      Else
         m_BillingDoc.Payments.Item(ID2).Flag = "D"
      End If

      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_BillingDoc.Payments)
      GridEX1.Rebind
      m_HasModify = True

   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   If Area = 1 Then
      If Not VerifyCombo(lblAccountNo, cboAccount) Then
         Exit Sub
      End If
      If Not VerifyDate(lblDocumentDate, uctlDocumentDate) Then
         Exit Sub
      End If
   End If
   
   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If Area = 1 Then
         frmAddEditDoItem.AccountID = cboAccount.ItemData(cboAccount.ListIndex)
      End If
      frmAddEditDoItem.DocumentDate = uctlDocumentDate.ShowDate
      frmAddEditDoItem.SubscriberID = -1
      frmAddEditDoItem.Area = Area
      frmAddEditDoItem.ID = ID
'      frmAddEditDoItem.DeliveryCostFlag = EditDeliveryCostFlag
      frmAddEditDoItem.COMMIT_FLAG = m_BillingDoc.OLD_COMMIT_FLAG
      Set frmAddEditDoItem.TempCollection = m_BillingDoc.DoItems
      frmAddEditDoItem.HeaderText = MapText("แก้ไขรายการใบส่งสินค้า")
      
      frmAddEditDoItem.ParentShowMode = ShowMode
      frmAddEditDoItem.ShowMode = SHOW_EDIT
      Load frmAddEditDoItem
      frmAddEditDoItem.Show 1

      OKClick = frmAddEditDoItem.OKClick

      Unload frmAddEditDoItem
      Set frmAddEditDoItem = Nothing
      
'      If EditDeliveryCostFlag Then
'         Dim tempDoItem As CDoItem
'         Set tempDoItem = m_BillingDoc.DoItems(ID)
'         If Not tempDoItem Is Nothing Then
'            txtNote.Text = tempDoItem.SUPPLIER_TRANSPORT_DETAIL
'         End If
'      End If

      If OKClick Then
         Call GetTotalPrice
         GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
         GridEX1.Rebind
'      Else
'         If EditDeliveryCostFlag Then
'            EditDeliveryCostFlag = False
'         End If
      End If
      
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If Area = 2 Then
         glbErrorLog.LocalErrorMsg = "ฟังก์ชันนี้ไม่สนับสนุนในส่วนงานซื้อ"
         glbErrorLog.ShowUserError
         Exit Sub
      End If
      frmAddEditBillingDiscount.ID = ID
      frmAddEditBillingDiscount.Area = Area
      frmAddEditBillingDiscount.COMMIT_FLAG = m_BillingDoc.COMMIT_FLAG
      Set frmAddEditBillingDiscount.TempCollection = m_BillingDoc.BillingDiscounts
      Set frmAddEditBillingDiscount.TempCollection2 = m_BillingDoc.DoItems
      frmAddEditBillingDiscount.ParentShowMode = ShowMode
      frmAddEditBillingDiscount.ShowMode = SHOW_EDIT
      frmAddEditBillingDiscount.HeaderText = MapText("แก้ไขรายการส่วนลด")
      Load frmAddEditBillingDiscount
      frmAddEditBillingDiscount.Show 1

      OKClick = frmAddEditBillingDiscount.OKClick

      Unload frmAddEditBillingDiscount
      Set frmAddEditBillingDiscount = Nothing

      If OKClick Then
         Call GetTotalPrice

         GridEX1.ItemCount = CountItem(m_BillingDoc.BillingDiscounts)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      frmAddEditCashTran.Area = Area
      Set frmAddEditCashTran.ParentForm = Me
      frmAddEditCashTran.ID = ID
      frmAddEditCashTran.HeaderText = "แก้ไขรายการการชำระเงิน"
      frmAddEditCashTran.ShowMode = SHOW_EDIT
      Set frmAddEditCashTran.TempCollection = m_BillingDoc.Payments
      Load frmAddEditCashTran
      frmAddEditCashTran.Show 1
      
      OKClick = frmAddEditCashTran.OKClick
      
      Unload frmAddEditCashTran
      Set frmAddEditCashTran = Nothing
   
      If OKClick Then
         m_HasModify = True
         
         GridEX1.ItemCount = CountItem(m_BillingDoc.Payments)
         Call GridEX1.Rebind
         
         Call GetTotalPrice
      End If

   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub



Private Sub cmdEditDeliveyCost_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim OKClick As Boolean
Dim TempCol As Collection
Dim Cs As CCustomer

   If Not VerifyAccessRight("LEDGER_SELL" & "_" & DocumentType & "_" & "TRANSPORT", "จัดการค่าขนส่ง") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
      
   Set TempCol = New Collection
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("เพิ่ม", "-", "แก้ไข", "-", "ลบ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
 If lMenuChosen = 1 Then
      frmAddEditTransport.DocumentDate = uctlDocumentDate.ShowDate
      frmAddEditTransport.DocumentNo = m_BillingDoc.DOCUMENT_NO
      frmAddEditTransport.BillingdocID = m_BillingDoc.BILLING_DOC_ID
      frmAddEditTransport.TruckNo = m_BillingDoc.NOTE
      frmAddEditTransport.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
      Set frmAddEditTransport.TempCollection = m_BillingDoc.DoItems
      Set frmAddEditTransport.TempCollection2 = m_BillingDoc.BillTransport
      frmAddEditTransport.HeaderText = MapText("เพิ่มรายการค่าขนส่งสินค้า")
     frmAddEditTransport.ShowMode = SHOW_ADD
      Load frmAddEditTransport
      frmAddEditTransport.Show 1

      OKClick = frmAddEditTransport.OKClick

      Unload frmAddEditTransport
      Set frmAddEditTransport = Nothing
      
      If OKClick Then
          m_BillingDoc.QueryFlag = 1
         Call QueryData(True)
      End If
   ElseIf lMenuChosen = 3 Then
      If CountItem(m_BillingDoc.BillTransport) = 0 Then
         glbErrorLog.LocalErrorMsg = "ยังไม่มีข้อมูลค่าขนส่ง"
         glbErrorLog.ShowUserError
         Exit Sub
      End If
      frmAddEditTransport.ID = ID
      frmAddEditTransport.DocumentDate = uctlDocumentDate.ShowDate
      frmAddEditTransport.DocumentNo = m_BillingDoc.DOCUMENT_NO
      frmAddEditTransport.BillingdocID = m_BillingDoc.BILLING_DOC_ID
      frmAddEditTransport.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
      Set frmAddEditTransport.TempCollection = m_BillingDoc.DoItems
      Set frmAddEditTransport.TempCollection2 = m_BillingDoc.BillTransport
      frmAddEditTransport.HeaderText = MapText("เพิ่มรายการค่าขนส่งสินค้า")
     frmAddEditTransport.ShowMode = SHOW_EDIT
      Load frmAddEditTransport
      frmAddEditTransport.Show 1

      OKClick = frmAddEditTransport.OKClick

      Unload frmAddEditTransport
      Set frmAddEditTransport = Nothing
      
      If OKClick Then
          m_BillingDoc.QueryFlag = 1
         Call QueryData(True)
      End If
   ElseIf lMenuChosen = 5 Then
     Dim m_BillTransport As CBillTransport
     Set m_BillTransport = New CBillTransport
     For Each m_BillTransport In m_BillingDoc.BillTransport
      Call m_BillTransport.DeleteData
     Next m_BillTransport

      m_BillingDoc.QueryFlag = 1
      Call QueryData(True)
   End If
'   If lMenuChosen = 1 Then
''      frmAddEditTransport.ID = ID
'      frmAddEditTransport.DocumentDate = uctlDocumentDate.ShowDate
'      frmAddEditTransport.DocumentNo = m_BillingDoc.DOCUMENT_NO
'      frmAddEditTransport.BillingdocID = m_BillingDoc.BILLING_DOC_ID
'      frmAddEditTransport.TruckNo = m_BillingDoc.NOTE
'      Set frmAddEditTransport.TempCollection = m_BillingDoc.DoItems
'      Set frmAddEditTransport.TempCollection2 = m_BillingDoc.BillTransport
'      frmAddEditTransport.HeaderText = MapText("เพิ่มรายการค่าขนส่งสินค้า")
'     frmAddEditTransport.ShowMode = SHOW_ADD
'      Load frmAddEditTransport
'      frmAddEditTransport.Show 1
'
'      OKClick = frmAddEditTransport.OKClick
'
'      Unload frmAddEditTransport
'      Set frmAddEditTransport = Nothing
'
'      If OKClick Then
'          m_BillingDoc.QueryFlag = 1
'         Call QueryData(True)
'      End If
'   ElseIf lMenuChosen = 3 Then
'      If CountItem(m_BillingDoc.BillTransport) = 0 Then
'         glbErrorLog.LocalErrorMsg = "ยังไม่มีข้อมูลค่าขนส่ง"
'         glbErrorLog.ShowUserError
'         Exit Sub
'      End If
'      frmAddEditTransport.id = id
'      frmAddEditTransport.DocumentDate = uctlDocumentDate.ShowDate
'      frmAddEditTransport.DocumentNo = m_BillingDoc.DOCUMENT_NO
'      frmAddEditTransport.BillingdocID = m_BillingDoc.BILLING_DOC_ID
'      Set frmAddEditTransport.TempCollection = m_BillingDoc.DoItems
'      Set frmAddEditTransport.TempCollection2 = m_BillingDoc.BillTransport
'      frmAddEditTransport.HeaderText = MapText("เพิ่มรายการค่าขนส่งสินค้า")
'     frmAddEditTransport.ShowMode = SHOW_EDIT
'      Load frmAddEditTransport
'      frmAddEditTransport.Show 1
'
'      OKClick = frmAddEditTransport.OKClick
'
'      Unload frmAddEditTransport
'      Set frmAddEditTransport = Nothing
'
'      If OKClick Then
'          m_BillingDoc.QueryFlag = 1
'         Call QueryData(True)
'      End If
'   ElseIf lMenuChosen = 5 Then
'     Dim m_BillTransport As CBillTransport
'     Set m_BillTransport = New CBillTransport
'     For Each m_BillTransport In m_BillingDoc.BillTransport
'      Call m_BillTransport.DeleteData
'     Next m_BillTransport
'
'      m_BillingDoc.QueryFlag = 1
'      Call QueryData(True)
'   End If
End Sub

Private Sub cmdOK_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen  As Long
Dim iCount As Long
Dim m_Rs1 As ADODB.Recordset
Dim TempError As String

   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("บันทึก", "-", "บันทึกและออกจากหน้าจอ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   'ตรวจเอกสาร
   If ID <> 0 Then
      Dim RT As CReceiptItem
      Set RT = New CReceiptItem
      Set m_Rs1 = New ADODB.Recordset
      RT.RECEIPT_ITEM_ID = -1
      RT.DO_ID = ID
      Call RT.QueryData(108, m_Rs1, iCount)
      Set RT = Nothing
      
      While Not m_Rs1.EOF
         Set RT = New CReceiptItem
         Call RT.PopulateFromRS(108, m_Rs1)
         
         If Len(Trim(TempError)) <= 0 Then
            TempError = RT.DOCUMENT_NO
         Else
            TempError = TempError & ", " & RT.DOCUMENT_NO
         End If
            
         Set RT = Nothing
         m_Rs1.MoveNext
      Wend
      
'      If Not EditDeliveryCostFlag Then
'         If Len(Trim(TempError)) > 0 Then
'            glbErrorLog.LocalErrorMsg = MapText("มีการอ้างอิงกับเอกสาร " & TempError & " โปรดลบเอกสารอ้างอิงดังกล่าวก่อน เพิ่มเติม/แก้ไข เอกสาร " & txtDocumentNo.Text)
'            glbErrorLog.ShowUserError
'            Exit Sub
'         End If
'      End If
   End If

   If lMenuChosen = 1 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      ShowMode = SHOW_EDIT
      ID = m_BillingDoc.BILLING_DOC_ID
      m_BillingDoc.QueryFlag = 1
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

Private Function VerifyOnwerVersionMenu(Menu As Long, Owner As String) As Boolean
   VerifyOnwerVersionMenu = True
   
   If (Menu <> 1) And (Menu <> 2) Then
      If Not glbGuiConfigs.VerifyGuiConfig("SELL_DO_PREFORM_PRINT", True) Then
         VerifyOnwerVersionMenu = False
         Exit Function
      End If
   End If
End Function

Private Sub cmdPrint_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim ReportFlag As Boolean
Dim ReportKey As String
Dim Report As CReportInterface
Dim Rc As CReportConfig
Dim iCount As Long
Dim EditMode As SHOW_MODE_TYPE
Dim ReportMode As Long

   ReportMode = 1
   
   If m_HasModify Or (m_BillingDoc.BILLING_DOC_ID <= 0) Then
      glbErrorLog.LocalErrorMsg = MapText("กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   ReportFlag = False

   Call LoadPictureFromFile(glbParameterObj.DOPicture1, Picture2)
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.AddMenu(glbGuiConfigs.DOPrintMenuItems)
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
'   If Not VerifyOnwerVersionMenu(lMenuChosen, glbParameterObj.Programowner) Then
'      Exit Sub
'   End If
   
   
   If lMenuChosen = 1 Then
      If m_BillingDoc.POST_FLAG = "N" Then
         glbErrorLog.LocalErrorMsg = MapText("เอกสารใบนี้ ยังไม่สมบูรณ์ ยังไม่สามารถพิมพ์ใบส่งของได้")
         glbErrorLog.ShowUserError
         Exit Sub
      End If
      ReportKey = "CReportNormalDO001"
      
      Set Report = New CReportNormalDO001
      ReportFlag = True
      Call Report.AddParam(False, "ExampleDoc")
      Call Report.AddParam(False, "PrintNotPrice")
      
   ElseIf lMenuChosen = 2 Then
      ReportKey = "CReportNormalDO001"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
     HeaderText = MapText("ใบส่งสินค้า/ใบแจ้งหนี้")
     
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
      
    ElseIf lMenuChosen = 10 Then
      ReportKey = "CReportNormalDoHead"
      
      Set Report = New CReportNormalDoHead
      ReportFlag = True
   ElseIf lMenuChosen = 11 Then
      ReportKey = "CReportNormalDoHead"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("INVOIVE")
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
      
   ElseIf lMenuChosen = 27 Then
      ReportKey = "CReportFormReceipt002"
      
      Set Report = New CReportFormReceipt002
      ReportFlag = True
   ElseIf lMenuChosen = 28 Then
      ReportMode = 2
      ReportKey = "CReportFormReceipt002"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
     HeaderText = MapText("แบบฟอร์มสยามธรรมาภิบาล")
     
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   ElseIf lMenuChosen = 30 Then
       ReportKey = "CReportNormalDO001"
      
      Set Report = New CReportNormalDO001
      
      Call Report.AddParam(True, "ExampleDoc")
      Call Report.AddParam(False, "PrintNotPrice")
      
      Picture2.Picture = LoadPicture(glbParameterObj.ExampleDoc)
      Call Report.AddParam(Picture2.Picture, "BACK_GROUND2")
      
      ReportFlag = True
   ElseIf lMenuChosen = 31 Then
      ReportKey = "CReportNormalDO001"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
     HeaderText = MapText("ใบส่งสินค้า/ใบแจ้งหนี้")
     
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
 ElseIf lMenuChosen = 33 Then
       If m_BillingDoc.POST_FLAG = "N" Then
         glbErrorLog.LocalErrorMsg = MapText("เอกสารใบนี้ ยังไม่สมบูรณ์ ยังไม่สามารถพิมพ์ใบส่งของได้")
         glbErrorLog.ShowUserError
         Exit Sub
      End If
       ReportKey = "CReportNormalDO001"
      
      Set Report = New CReportNormalDO001
      
      Call Report.AddParam(False, "ExampleDoc")
      Call Report.AddParam(True, "PrintNotPrice")

      ReportFlag = True
   ElseIf lMenuChosen = 34 Then
      ReportKey = "CReportNormalDO001"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
     HeaderText = MapText("ใบส่งสินค้า/ใบแจ้งหนี้")
     
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   End If

   If Not Report Is Nothing Then
      Call Report.AddParam(lMenuChosen, "REPORT_TYPE")
      Call Report.AddParam(m_BillingDoc.BILLING_DOC_ID, "BILLING_DOC_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      Call Report.AddParam(MapText("ใบส่งสินค้า/ใบแจ้งหนี้"), "REPORT_HEADER")
      Call Report.AddParam(DocumentType, "DOCUMENT_TYPE")
      
      Dim C As Long
      C = 0
      If lMenuChosen = 1 Or lMenuChosen = 2 Then
         C = m_BillingDoc.getPrintCount(m_BillingDoc.BILLING_DOC_ID)
         If C > 0 Then
           lblMsg.Caption = MapText("เอกสารใบนี้ ได้ทำการพิมพ์ไป " & C & " ครั้งแล้ว")
   '      glbErrorLog.LocalErrorMsg = MapText("เอกสารใบนี้ ได้ทำการพิมพ์ไป " & C & " ครั้งแล้ว")
   '      glbErrorLog.ShowUserError
         End If
         glbParameterObj.PrintCount = C
      End If
      
      glbParameterObj.ReportKey = ReportKey
      glbParameterObj.ID = m_BillingDoc.BILLING_DOC_ID
      glbParameterObj.DocType = DocumentType
      glbParameterObj.PrintCount = C
      
      Call Report.AddParam(Picture2.Picture, "BACK_GROUND")
      Call Report.AddParam(uctlSellByLookup.MyCombo.Text, "RECEIVE_NAME")
      Call Report.AddParam("", "ACCEPT_NAME")
   ElseIf lMenuChosen = 5 Then
      ReportKey = "CReportFormPO001"
      
      Set Rc = New CReportConfig
      Rc.REPORT_KEY = ReportKey
      Rc.COMPUTER_NAME = glbDatabaseMngr.GetComputerName
      Call Rc.QueryData(m_Rs, iCount)
      HeaderText = MapText("ใบส่งสินค้า/ใบแจ้งหนี้")
      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   End If
   
   Call EnableForm(Me, False)
   If ReportFlag Then
      Set frmReport.ReportObject = Report
      frmReport.HeaderText = ""
      Load frmReport
      frmReport.Show 1
         
      Unload frmReport
      Set frmReport = Nothing
      Set Report = Nothing
   Else
      frmReportConfig.ReportMode = ReportMode
      frmReportConfig.ShowMode = EditMode
      frmReportConfig.ID = Rc.REPORT_CONFIG_ID
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
   ID = m_BillingDoc.BILLING_DOC_ID
   m_BillingDoc.QueryFlag = 1
   QueryData (True)
   m_HasModify = False
End Sub

Private Sub LoadDoPartItem()
Dim Di As CDoItem
Dim I As Long

   cboHold1.Clear
   cboHold2.Clear
   cboHold3.Clear
   cboHold4.Clear
   
   cboHold1.AddItem ("")
   cboHold2.AddItem ("")
   cboHold3.AddItem ("")
   cboHold4.AddItem ("")
   
   I = 0
   For Each Di In m_BillingDoc.DoItems
      If (Di.Flag <> "D") And (Di.PART_ITEM_ID > 0) Then
         I = I + 1
         cboHold1.AddItem (Di.PART_NO)
         cboHold1.ItemData(I) = Di.PART_ITEM_ID
      
         cboHold2.AddItem (Di.PART_NO)
         cboHold2.ItemData(I) = Di.PART_ITEM_ID
      
         cboHold3.AddItem (Di.PART_NO)
         cboHold3.ItemData(I) = Di.PART_ITEM_ID
         
         cboHold4.AddItem (Di.PART_NO)
         cboHold4.ItemData(I) = Di.PART_ITEM_ID
      End If
   Next Di
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
'      DoEvents
      
      AllowSave = True
      
      Call EnableForm(Me, False)
      Call LoadEnterpriseAddress(cboEnpAddress, , , True)
      
      Call InitPaymentType(cboPaymentType)
      Call LoadBank(cboBank)
      
      Call InitPackageType(cboPackageType)
      Call LoadConfigDoc(Nothing, m_Cd)
      
      
      If DocumentType = 1 Then
         Call LoadDistinctCustomerPicture(Nothing, m_CustomerPictures)
      End If
      
      If Area = 1 Then
         Call LoadCustomer(uctlCustomerLookup.MyCombo, m_Customers)
         Set uctlCustomerLookup.MyCollection = m_Customers
      ElseIf Area = 2 Then
         Call LoadSupplier(uctlCustomerLookup.MyCombo, m_Customers)
         Set uctlCustomerLookup.MyCollection = m_Customers
      End If

      Call LoadEmployee(uctlSellByLookup.MyCombo, m_Employees)
      Set uctlSellByLookup.MyCollection = m_Employees

      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_BillingDoc.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
         
      ElseIf ShowMode = SHOW_ADD Then
         Call LoadDoPartItem
         
         uctlDocumentDate.ShowDate = Now
         uctlShipDate.ShowDate = Now
         m_BillingDoc.QueryFlag = 0
         Call QueryData(False)
         
      End If
      DocAdd = 0
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Static InUsed As Long

   If InUsed = 1 Then
      Exit Sub
   End If
   
   InUsed = 1
   
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
   
   InUsed = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_BillingDoc = Nothing
   Set m_Customers = Nothing
   Set m_Employees = Nothing
   Set m_Resources = Nothing
   Set m_CustomerPictures = Nothing
   Set m_Cd = Nothing
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

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 2325 + 2055 + 2235
   Col.Caption = MapText("รายละเอียด")

   Set Col = GridEX1.Columns.add '4
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1620
   Col.Caption = MapText("จำนวน")
      
   Set Col = GridEX1.Columns.add '5
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1755
   Col.Caption = MapText("ราคา/หน่วย")
   
   Set Col = GridEX1.Columns.add '6
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1575
   Col.Caption = MapText("ราคารวม")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = 2325
   Col.Caption = MapText("เลขที่ PO")
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
   Col.Width = 2805
   Col.Caption = MapText("ชื่อส่วนลด")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 5055 + 1950
   Col.Caption = MapText("ชื่อสินค้า")
   
   Set Col = GridEX1.Columns.add '5
   Col.TextAlignment = jgexAlignRight
   Col.Width = 1755
   Col.Caption = MapText("มูลค่าส่วนลด")
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

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

      Set Col = GridEX1.Columns.add '3
      Col.Width = 1965
      Col.Caption = MapText("ประเภทการชำระเงิน")
   
      Set Col = GridEX1.Columns.add '4
      Col.Width = 2625
      Col.Caption = MapText("เลขที่เช็ค/บัญชี")
   
      Set Col = GridEX1.Columns.add '5
      Col.Width = 2160
      Col.TextAlignment = jgexAlignLeft
      Col.Caption = MapText("ธนาคาร")
   
      Set Col = GridEX1.Columns.add '6
      Col.Width = 2565
      Col.TextAlignment = jgexAlignLeft
      Col.Caption = MapText("สาขาธนาคาร")
   
      Set Col = GridEX1.Columns.add '7
      Col.Width = 2000
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("จำนวนเงิน")

End Sub
Private Sub GetTotalPrice()
Dim II As CDoItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim Sum4 As Double
Dim Sum5 As Double
Dim Sum6 As Double
Dim Bds As CBillingDiscount
Dim Sum7 As Double
Dim Pm As CCashTran

   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   Sum4 = 0
   Sum5 = 0

   For Each II In m_BillingDoc.DoItems
      If II.Flag <> "D" Then
         Sum1 = Sum1 + II.ITEM_AMOUNT
         Sum2 = Sum2 + (II.TOTAL_PRICE + II.DISCOUNT_AMOUNT)
         Sum3 = Sum3 + II.TOTAL_WEIGHT
         Sum4 = Sum4 + II.DISCOUNT_AMOUNT
         Sum5 = Sum5 + II.DEPOSIT_AMOUNT
      End If
   Next II

   Sum6 = 0
   For Each Bds In m_BillingDoc.BillingDiscounts
      If Bds.Flag <> "D" Then
         Sum6 = Sum6 + Bds.DISCOUNT_AMOUNT
      End If
   Next Bds
   
   Sum7 = 0
   For Each Pm In m_BillingDoc.Payments
      Sum7 = Sum7 + Pm.GetFieldValue("AMOUNT")
   Next Pm
   
   txtNetTotal.Text = Format(Sum2, "0.00")
'   txtTotalDiscount.Text = Format(Sum3, "0.00")
   txtTotalAmount.Text = Format(Sum1, "0.00")
   txtDiscount.Text = Format(Sum4, "0.00")
   txtCashDiscountAmount.Text = Format(Sum6, "0.00")
   txtTotalRcp.Text = Format(Sum7, "0.00")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame3.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame4.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Programowner = glbParameterObj.Programowner
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่ใบส่งสินค้า"))
   Call InitNormalLabel(lblAccountNo, MapText("เลขที่บัญชี"))
   If Area = 1 Then
      Call InitNormalLabel(lblCustomerAddress, MapText("ที่อยู่ลูกค้า"))
      Call InitNormalLabel(lblCustomer, MapText("รหัสลูกค้า"))
      Call InitNormalLabel(lblSellBy, MapText("พนักงานขาย"))
      Call InitNormalLabel(lblEnpAddress, MapText("ที่อยู่ออกเอกสาร"))
   ElseIf Area = 2 Then
      Call InitNormalLabel(lblCustomerAddress, MapText("ที่อยู่ซัพ ฯ"))
      Call InitNormalLabel(lblCustomer, MapText("รหัสซัพ ฯ"))
      Call InitNormalLabel(lblSellBy, MapText("ผู้รับของ"))
      Call InitNormalLabel(lblEnpAddress, MapText("ที่อยู่รับเอกสาร"))
      
      lblAccountNo.Visible = False
      cboAccount.Visible = False
      cmdAuto.Visible = False
      cmdCustomer.Visible = False
      cmdPrint.Enabled = False
   End If
   Call InitNormalLabel(lblTotalAmount, MapText("จำนวนรวม"))
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่เอกสาร"))
   Call InitNormalLabel(Label1, MapText("ตัว"))
   Call InitNormalLabel(Label4, MapText("บาท"))
   Call InitNormalLabel(Label5, MapText("%"))
   Call InitNormalLabel(lblNetTotal, MapText("ราคารวม"))
   Call InitNormalLabel(lblDiscount, MapText("ส่วนลด"))
   Call InitNormalLabel(lblCashDiscount, MapText("% ส่วนลดเงินสด"))
   Call InitNormalLabel(lblIncludeDiscount, MapText("รวมส่วนลด"))
   Call InitNormalLabel(lblLeft, MapText("คงค้าง"))
   Call InitCheckBox(chkCommit, "คำนวณ")
   Call InitCheckBox(chkPostFlag, "เอกสารสมบูรณ์")
   Call InitNormalLabel(Label3, MapText("บาท"))
   Call InitNormalLabel(Label6, MapText("บาท"))
   Call InitNormalLabel(Label8, MapText("บาท"))
   Call InitNormalLabel(Label10, MapText("บาท"))
   Call InitNormalLabel(lblDueDate, MapText("วันนัดชำระ"))
   Call InitNormalLabel(lblNote, MapText("ทะเบียนรถ"))
   Call InitNormalLabel(lblShipment, MapText("วันที่ส่งของ"))
   Call InitNormalLabel(lblDeliveryPlace, MapText("สถานที่จัดส่ง"))
   Call InitNormalLabel(lblPoNo, MapText("เลขที่ใบสั่งซื้อ"))
   Call InitNormalLabel(Label7, MapText("ส่วนลดเพิ่มเติม"))
   Call InitNormalLabel(Label9, MapText("บาท"))
   Call InitNormalLabel(Label2, MapText("บาท"))
   Call InitNormalLabel(Label15, MapText("บาท"))
   Call InitNormalLabel(Label11, MapText("วัน"))
   Call InitNormalLabel(lblCredit, MapText("เครดิต"))
   Call InitNormalLabel(lblPackageType, MapText("บรรจุ"))
   Call InitNormalLabel(lblInOutTime, MapText("เวลาเข้า-ออก"))
   Call InitNormalLabel(lblGeneration, MapText("รุ่น"))
   Call InitNormalLabel(lblReference, MapText("อ้างอิง"))
   Call InitNormalLabel(lblFarmName, MapText("ชื่อฟาร์ม"))
   Call InitNormalLabel(Label13, MapText("-"))
   
   Call InitNormalLabel(lblHold1, MapText("ช่อง 1"))
   Call InitNormalLabel(lblHold2, MapText("ช่อง 2"))
   Call InitNormalLabel(lblHold3, MapText("ช่อง 3"))
   Call InitNormalLabel(lblHold4, MapText("ช่อง 4"))
   Call InitNormalLabel(lblHold1Amount, MapText("จำนวน"))
   Call InitNormalLabel(lblHold2Amount, MapText("จำนวน"))
   Call InitNormalLabel(lblHold3Amount, MapText("จำนวน"))
   Call InitNormalLabel(lblHold4Amount, MapText("จำนวน"))
   Call InitNormalLabel(lblHold1Desc, MapText("หมายเหตุ"))
   Call InitNormalLabel(lblHold2Desc, MapText("หมายเหตุ"))
   Call InitNormalLabel(lblHold3Desc, MapText("หมายเหตุ"))
   Call InitNormalLabel(lblHold4Desc, MapText("หมายเหตุ"))
   Call InitNormalLabel(lblTempDONo, MapText("ใบส่งของชั่วคราว"))
   
   Call InitNormalLabel(lblPaymentType, MapText("การชำระเงิน"))
   Call InitNormalLabel(lblCheckNo, MapText("เลขที่เช็ค"))
   Call InitNormalLabel(lblCheckDate, MapText("วันที่เช็ค"))
   Call InitNormalLabel(lblBank, MapText("ธนาคาร"))
   Call InitNormalLabel(lblBankBranch, MapText("สาขา"))
   
   Call InitNormalLabel(lblTotalRcp, MapText("ยอดชำระจริง"))
   Call InitNormalLabel(lblDipRcp, MapText("ผลต่างรับชำระ"))
   
   Call InitCombo(cboPaymentType)
   Call InitCombo(cboBank)
   Call InitCombo(cboBankBranch)
   
   Call txtCheckNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Call txtPayment.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtNote.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtTotalAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalAmount.Enabled = False
'   Call txtTotalDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
'   txtTotalDiscount.Enabled = False
   Call txtNetTotal.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtNetTotal.Enabled = False
   Call txtCashDiscount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtIncludeDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtIncludeDiscount.Enabled = False
   Call txtDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtDiscount.Enabled = False
   Call txtLeft.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtLeft.Enabled = False
   Call txtPONo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtCashDiscountAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtCashDiscountAmount.Enabled = False
   Call txtCredit.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   Call txtTempDONo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Call txtTotalRcp.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtDipRcp.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalRcp.Enabled = False
   txtDipRcp.Enabled = False
   
   Call txtHold1Amount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtHold1Desc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtHold2Amount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtHold2Desc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtHold3Amount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtHold3Desc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtHold4Amount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtHold4Desc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   If DocumentType = 1 Then
     If Not VerifyAccessRight("LEDGER_SELL" & "_" & DocumentType & "_" & "POST", "กำหนดเอกสารสมบูรณ์", 2) Then
        chkPostFlag.Enabled = False
      Else
        chkPostFlag.Enabled = True
     End If
   End If
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   GridEX1.Visible = True
   SSFrame2.Visible = False
   SSFrame3.Visible = False
   SSFrame4.Visible = False
   
   Call InitCombo(cboAccount)
   Call InitCombo(cboCustomerAddress)
   Call InitCombo(cboEnpAddress)
   Call InitCombo(cboPackageType)
   
   Call InitCombo(cboHold1)
   Call InitCombo(cboHold2)
   Call InitCombo(cboHold3)
   Call InitCombo(cboHold4)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdCustomer.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEditDeliveyCost.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdCustomer, MapText("F"))
   Call InitMainButton(cmdEditDeliveyCost, MapText("ค่าขนส่ง"))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.NAME = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("รายการใบส่งสินค้า")
   TabStrip1.Tabs.add().Caption = MapText("ส่วนลดเพิ่มเติม")
   TabStrip1.Tabs.add().Caption = MapText("รายละเอียดทั่วไป")
   TabStrip1.Tabs.add().Caption = MapText("ช่อง Bulk")
   If DocumentType = 2 Then
      TabStrip1.Tabs.add().Caption = MapText("การชำระเงิน")
   End If
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
   Set m_BillingDoc = New CBillingDoc
   Set m_Customers = New Collection
   Set m_Employees = New Collection
   Set m_Resources = New Collection
   Set m_CustomerPictures = New Collection
   Set m_Cd = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

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

   glbErrorLog.ModuleName = Me.NAME
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_BillingDoc.DoItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CDoItem
      If m_BillingDoc.DoItems.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_BillingDoc.DoItems, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.DO_ITEM_ID
      Values(2) = RealIndex
      Values(3) = CR.ShowDescText
      Values(4) = FormatNumber(CR.ITEM_AMOUNT)
      Values(5) = FormatNumber(CR.AVG_PRICE)
      Values(6) = FormatNumber(CR.TOTAL_PRICE)
      Values(7) = CR.PO_NO
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If m_BillingDoc.BillingDiscounts Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Bdsc As CBillingDiscount
      If m_BillingDoc.BillingDiscounts.Count <= 0 Then
         Exit Sub
      End If
      Set Bdsc = GetItem(m_BillingDoc.BillingDiscounts, RowIndex, RealIndex)
      If Bdsc Is Nothing Then
         Exit Sub
      End If

      Values(1) = Bdsc.BILLING_DISCOUNT_ID
      Values(2) = RealIndex
      Values(3) = Bdsc.DISCOUNT_NAME
      If Bdsc.FEATURE_ID > 0 Then
         Values(4) = Bdsc.ITEM_DESC
      ElseIf Bdsc.PART_ITEM_ID > 0 Then
         Values(4) = Bdsc.ITEM_DESC
      End If
      Values(5) = FormatNumber(Bdsc.DISCOUNT_AMOUNT)
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      If m_BillingDoc.Payments Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Ct As CCashTran
      If m_BillingDoc.Payments.Count <= 0 Then
         Exit Sub
      End If
      Set Ct = GetItem(m_BillingDoc.Payments, RowIndex, RealIndex)
      If Ct Is Nothing Then
         Exit Sub
      End If
      
      Values(1) = Ct.GetFieldValue("CASH_TRAN_ID")
      Values(2) = RealIndex
      Values(3) = Ct.GetFieldValue("PAYMENT_TYPE_NAME")
      If Ct.GetFieldValue("PAYMENT_TYPE") = CASH_PMT Then
         Values(7) = FormatNumber(Ct.GetFieldValue("AMOUNT"))
      ElseIf Ct.GetFieldValue("PAYMENT_TYPE") = CREDITCRD_PMT Then
         Values(4) = Ct.GetFieldValue("ACCOUNT_NAME")
         Values(5) = Ct.GetFieldValue("BANK_NAME")
         Values(6) = Ct.GetFieldValue("BRANCH_NAME")
         Values(7) = FormatNumber(Ct.GetFieldValue("AMOUNT"))
      ElseIf Ct.GetFieldValue("PAYMENT_TYPE") = CHECK_PMT Then
         Values(4) = Ct.Cheque.GetFieldValue("CHEQUE_NO")
         Values(5) = Ct.Cheque.GetFieldValue("BANK_NAME")
         Values(6) = Ct.Cheque.GetFieldValue("BRANCH_NAME")
         Values(7) = FormatNumber(Ct.GetFieldValue("AMOUNT"))
      End If
      
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub EnableDisableButton(En As Boolean)
   If En Then
      If ShowMode = SHOW_EDIT Then
         cmdAdd.Enabled = (m_BillingDoc.OLD_COMMIT_FLAG = "N")
         cmdEdit.Enabled = True '(m_BillingDoc.COMMIT_FLAG = "N")
         cmdDelete.Enabled = (m_BillingDoc.OLD_COMMIT_FLAG = "N")
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






Private Sub TabStrip1_Click()
   GridEX1.Top = 5160
   GridEX1.Left = 150
   GridEX1.Visible = False
   
   SSFrame2.Top = 5160
   SSFrame2.Left = 150
   SSFrame2.Visible = False
   
   SSFrame3.Top = 5160
   SSFrame3.Left = 150
   SSFrame3.Visible = False
   
   SSFrame4.Top = 5160
   SSFrame4.Left = 150
   SSFrame4.Visible = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      Call EnableDisableButton(True)
      Call InitGrid1
      GridEX1.Visible = True
      
      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_BillingDoc.DoItems)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      Call EnableDisableButton(True)
      Call InitGrid2
      GridEX1.Visible = True
      
      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_BillingDoc.BillingDiscounts)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      Call EnableDisableButton(False)
      SSFrame2.Visible = True
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      Call EnableDisableButton(False)
      SSFrame3.Visible = True
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
'      Call EnableDisableButton(False)
'      SSFrame4.Visible = True
      Call EnableDisableButton(True)
      Call InitGrid3
      GridEX1.Visible = True

      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_BillingDoc.Payments)
      GridEX1.Rebind
   End If
End Sub

Private Sub txtCashDiscount_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtCashDiscountAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtCheckNo_Change()
   m_HasModify = True
End Sub

Private Sub txtCredit_Change()
Dim NewDate As Date

   m_HasModify = True

   NewDate = DateAdd("D", Val(txtCredit.Text), uctlDocumentDate.ShowDate)
   uctlDueDate.ShowDate = NewDate
End Sub

Private Sub txtDiscount_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtFarmName_Change()
m_HasModify = True
End Sub

Private Sub txtGeneration_Change()
      m_HasModify = True
End Sub

Private Sub txtHold1Amount_Change()
   m_HasModify = True
End Sub

Private Sub txtHold1Desc_Change()
   m_HasModify = True
End Sub

Private Sub txtHold2Amount_Change()
   m_HasModify = True
End Sub

Private Sub txtHold2Desc_Change()
   m_HasModify = True
End Sub

Private Sub txtHold3Amount_Change()
   m_HasModify = True
End Sub

Private Sub txtHold3Desc_Change()
   m_HasModify = True
End Sub

Private Sub txtHold4Amount_Change()
   m_HasModify = True
End Sub

Private Sub txtHold4Desc_Change()
   m_HasModify = True
End Sub

Private Sub txtIncludeDiscount_Change()
   m_HasModify = True
   Call CalculateAmount
End Sub

Private Sub txtLeft_Change()
   m_HasModify = True
End Sub

Private Sub txtNetTotal_Change()
   Call CalculateAmount
   m_HasModify = True
End Sub

Private Sub txtNote_Change()
m_HasModify = True
End Sub

Private Sub txtPayment_Change()
m_HasModify = True
End Sub

Private Sub txtPONo_Change()
   m_HasModify = True
End Sub

Private Sub txtReference_Change()
m_HasModify = True
End Sub

Private Sub txtTempDONo_Change()
   m_HasModify = True
End Sub

Private Sub txtTotalAmount_Change()
   m_HasModify = True
End Sub

Private Sub CalculateAmount()
Dim TempAmt As Double

   txtIncludeDiscount.Text = Format(Val(txtNetTotal.Text) - Val(txtDiscount.Text), "0.00")
   TempAmt = Val(txtIncludeDiscount.Text) * Val(txtCashDiscount.Text) / 100
   txtLeft.Text = Format(Val(txtIncludeDiscount.Text) - TempAmt - Val(txtCashDiscountAmount.Text), "0.00")
   txtDipRcp.Text = Format(Val(txtLeft.Text) - Val(txtTotalRcp.Text), "0.00")
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

'Private Sub txtTotalDiscount_Change()
'   m_HasModify = True
'   txtNetTotal.Text = Format(Val(txtTotalAmount.Text) + Val(txtTotalDiscount.Text), "0.00")
'End Sub

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

Private Sub txtTotalRcp_Change()
   Call CalculateAmount
   m_HasModify = True
End Sub

Private Sub uctlCheckDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlCustomerLookup_Change()
Dim CustomerID As Long
Dim Customer As CCustomer
Static OldCusId As Long

   CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   If CustomerID > 0 Then
      If OldCusId = CustomerID Then
         Exit Sub
      Else
         OldCusId = CustomerID
      End If
      
      If Area = 1 Then
         Set Customer = m_Customers(Trim(str(CustomerID)))
         Call LoadAccount(cboAccount, , CustomerID)
         cboAccount.ListIndex = 1
         
         Call LoadCustomerAddress(cboCustomerAddress, , CustomerID, True)
         txtCredit.Text = Customer.Credit
         If Customer.RESPONSE_BY > 0 Then
            uctlSellByLookup.MyCombo.ListIndex = IDToListIndex(uctlSellByLookup.MyCombo, Customer.RESPONSE_BY)
         Else
            uctlSellByLookup.MyCombo.ListIndex = -1
         End If
      ElseIf Area = 2 Then
         Call LoadAccount(cboAccount, , CustomerID)
         cboAccount.ListIndex = -1
   
         Call LoadSupplierAddress(cboCustomerAddress, , CustomerID, True)
      End If
   Else
      cboAccount.ListIndex = -1
      cboCustomerAddress.ListIndex = -1
   End If
   m_HasModify = True
   
   If DocumentType = 1 And CustomerID > 0 Then
      Dim CP As CCustomerPicture
      Set CP = GetObject("CCustomerPicture", m_CustomerPictures, Trim(str(CustomerID)), False)
      If CP Is Nothing And DateDiff("D", Customer.CREATE_DATE, Now) > 15 Then
         glbErrorLog.LocalErrorMsg = MapText("ยังไม่มีข้อมูล ใบปะหน้าบัญชี และเวลาได้เกินกำหนดแล้วไม่สามารถสร้างเอกสารใบขายเชื่อได้ " & vbCrLf & "  กรุณาเพิ่มข้อมูลใบปะหน้าบัญชีเพิ่มให้สามารถใช้ Funtion นี้ได้")
         glbErrorLog.ShowUserError
         AllowSave = False
      ElseIf CP Is Nothing And DateDiff("D", Customer.CREATE_DATE, Now) <= 15 Then
         glbErrorLog.LocalErrorMsg = MapText("ยังไม่มีข้อมูล ใบปะหน้าบัญชี ท่านสามารถบันทึกใบขายเชื่อ ได้จนถึงวันที่ " & DateToStringExtEx2(DateAdd("D", 15, Customer.CREATE_DATE)) & vbCrLf & "  กรุณาเพิ่มข้อมูลใบปะหน้าบัญชีเพิ่มให้สามารถใช้ Funtion นี้ได้ต่อไป")
         glbErrorLog.ShowUserError
         AllowSave = True
      Else
         AllowSave = True
      End If
      Set CP = Nothing
   End If
   Set Customer = Nothing
End Sub

Private Sub uctlDueDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlPaymentDate_HasChange()
m_HasModify = True
End Sub

Private Sub uctlResource_Change()
   m_HasModify = True
End Sub

Private Sub uctlSellByLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlShipDate_HasChange()
m_HasModify = True
End Sub
Private Function CheckCreditLimit() As Double
Dim Doc As CDoItem
Dim Rcp As CReceiptItem
Dim Cn As CReceiptItem
Dim Dn As CReceiptItem
Dim RT As CReceiptItem
Dim BLD As CBillingDiscount
Dim m_Rs  As ADODB.Recordset
Dim ItemCount As Long
   
   CheckCreditLimit = 0
   
   Set Doc = New CDoItem
   Set m_Rs = New ADODB.Recordset
   Doc.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   Doc.EXCEPT_BILLING_ID = ID
   Call Doc.QueryData(9, m_Rs, ItemCount)
   If ItemCount > 0 Then
      Call Doc.PopulateFromRS(9, m_Rs)
   End If
   CheckCreditLimit = CheckCreditLimit + Doc.TOTAL_PRICE
   Set Doc = Nothing
   Set m_Rs = Nothing
   
   Set Rcp = New CReceiptItem
   Set m_Rs = New ADODB.Recordset
   Rcp.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   Call Rcp.QueryData(3, m_Rs, ItemCount)
   If ItemCount > 0 Then
      Call Rcp.PopulateFromRS(3, m_Rs)
   End If
   CheckCreditLimit = CheckCreditLimit - Rcp.PAID_AMOUNT - Rcp.CASH_DISCOUNT
   Set Rcp = Nothing
   Set m_Rs = Nothing
   
   Set Dn = New CReceiptItem
   Set m_Rs = New ADODB.Recordset
   Dn.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   Dn.DOCUMENT_TYPE = 4
   Call Dn.QueryData(6, m_Rs, ItemCount)
   If ItemCount > 0 Then
      Call Dn.PopulateFromRS(6, m_Rs)
   End If
   CheckCreditLimit = CheckCreditLimit + Dn.DEBIT_CREDIT_AMOUNT
   Set Dn = Nothing
   Set m_Rs = Nothing
   
   Set Cn = New CReceiptItem
   Set m_Rs = New ADODB.Recordset
   Cn.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   Cn.DOCUMENT_TYPE = 3
   Call Cn.QueryData(6, m_Rs, ItemCount)
   If ItemCount > 0 Then
      Call Cn.PopulateFromRS(6, m_Rs)
   End If
   CheckCreditLimit = CheckCreditLimit - Cn.DEBIT_CREDIT_AMOUNT
   Set Cn = Nothing
   Set m_Rs = Nothing
   
   Set RT = New CReceiptItem
   Set m_Rs = New ADODB.Recordset
   RT.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   RT.DOCUMENT_TYPE = 18
   Call RT.QueryData(6, m_Rs, ItemCount)
   If ItemCount > 0 Then
      Call RT.PopulateFromRS(6, m_Rs)
   End If
   CheckCreditLimit = CheckCreditLimit - RT.DEBIT_CREDIT_AMOUNT
   Set RT = Nothing
   Set m_Rs = Nothing
   
   Set BLD = New CBillingDiscount
   Set m_Rs = New ADODB.Recordset
   BLD.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   Call BLD.QueryData(4, m_Rs, ItemCount)
   If ItemCount > 0 Then
      Call BLD.PopulateFromRS(4, m_Rs)
   End If
   CheckCreditLimit = CheckCreditLimit - BLD.DISCOUNT_AMOUNT
   Set BLD = Nothing
   Set m_Rs = Nothing
   
End Function
Private Function CheckMaxCreditDueDate(Cs As CCustomer) As Boolean
Dim BD As CBillingDoc
Dim CustomerID As Long
Dim Rcp As CReceiptItem
Dim Cn As CReceiptItem
Dim Dn As CReceiptItem
Dim RT As CReceiptItem

Dim m_Rs  As ADODB.Recordset
Dim ItemCount As Long
Dim m_PaidAmounts  As Collection
Dim m_DnItemsByBill  As Collection
Dim m_CnItemsByBill  As Collection
Dim m_RtItemsByBill  As Collection
Dim m_BillingDiscounts As Collection
Dim Bdc As CBillingDiscount
Dim ServerDateTime As String
   CheckMaxCreditDueDate = True
      
   Set m_PaidAmounts = New Collection
   Set m_DnItemsByBill = New Collection
   Set m_CnItemsByBill = New Collection
   Set m_RtItemsByBill = New Collection
   Set m_BillingDiscounts = New Collection
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   
   Call LoadPaidAmountByCustomer2(m_PaidAmounts, -1, -1, CustomerID)
   Call LoadDnCnAmountByCustomer2(m_DnItemsByBill, -1, -1, 4, 2, CustomerID)
   Call LoadDnCnAmountByCustomer2(m_CnItemsByBill, -1, -1, 3, 2, CustomerID)
   Call LoadDnCnAmountByCustomer2(m_RtItemsByBill, -1, -1, 18, 2, CustomerID)
   Call LoadBillingDiscountByBill(Nothing, m_BillingDiscounts)
   
   Set BD = New CBillingDoc
   Set m_Rs = New ADODB.Recordset
   BD.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   BD.DOCUMENT_TYPE = 1
   Call BD.QueryData(103, m_Rs, ItemCount)
   

   While Not m_Rs.EOF
      Call BD.PopulateFromRS(103, m_Rs)
'      If Bd.BILLING_DOC_ID = 184776 Then
'         'Debug.Print
'      End If
'      ''Debug.Print ((DateDiff("D", Bd.DUE_DATE, InternalDateToDateEx2(ServerDateTime))))
      If ((DateDiff("D", BD.DUE_DATE, InternalDateToDateEx2(ServerDateTime)))) > Cs.MAX_CREDIT Then                       'เงิน Due ของบิลนั้นๆได้กี่วัน
         Set Rcp = GetObject("CReceiptItem", m_PaidAmounts, Trim(str(BD.BILLING_DOC_ID)))
         Set Cn = GetObject("CReceiptItem", m_CnItemsByBill, Trim(str(BD.BILLING_DOC_ID)))
         Set Dn = GetObject("CReceiptItem", m_DnItemsByBill, Trim(str(BD.BILLING_DOC_ID)))
         Set RT = GetObject("CReceiptItem", m_RtItemsByBill, Trim(str(BD.BILLING_DOC_ID)))
         Set Bdc = GetBillingDiscount(m_BillingDiscounts, Trim(str(BD.BILLING_DOC_ID)))
         
         If ROUND(BD.TOTAL_PRICE - BD.DISCOUNT_AMOUNT - Bdc.DISCOUNT_AMOUNT + Dn.DEBIT_CREDIT_AMOUNT - Cn.DEBIT_CREDIT_AMOUNT - RT.DEBIT_CREDIT_AMOUNT - Rcp.PAID_AMOUNT, 2) > 0 Then
            CheckMaxCreditDueDate = False
            Exit Function
         End If
      End If
      m_Rs.MoveNext
   Wend
   
   Set BD = Nothing
   Set m_Rs = Nothing
   
   Set m_PaidAmounts = Nothing
   Set m_DnItemsByBill = Nothing
   Set m_CnItemsByBill = Nothing
   Set m_RtItemsByBill = Nothing
End Function

Private Function CheckWeekCreditLimit(DocDate As Date) As Double
Dim Doc As CDoItem
Dim BLD As CBillingDiscount
Dim m_Rs  As ADODB.Recordset
Dim ItemCount As Long
Dim FromDate As Date
Dim ToDate As Date
Dim TempWeekDate As Long
   TempWeekDate = Weekday(DocDate, vbMonday)
   FromDate = DateAdd("D", 1 - TempWeekDate, DocDate)
   ToDate = DateAdd("D", 6, FromDate)
      
   
   CheckWeekCreditLimit = 0
   
   Set Doc = New CDoItem
   Set m_Rs = New ADODB.Recordset
   Doc.FROM_DATE = FromDate
   Doc.TO_DATE = ToDate
   Doc.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   Doc.EXCEPT_BILLING_ID = ID
   Call Doc.QueryData(9, m_Rs, ItemCount)
   If ItemCount > 0 Then
      Call Doc.PopulateFromRS(9, m_Rs)
   End If
   CheckWeekCreditLimit = CheckWeekCreditLimit + Doc.TOTAL_PRICE
   Set m_Rs = Nothing
   
   Set BLD = New CBillingDiscount
   Set m_Rs = New ADODB.Recordset
   BLD.FROM_DATE = FromDate
   BLD.TO_DATE = ToDate
   BLD.CUSTOMER_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   Call BLD.QueryData(4, m_Rs, ItemCount)
   If ItemCount > 0 Then
      Call BLD.PopulateFromRS(4, m_Rs)
   End If
   CheckWeekCreditLimit = CheckWeekCreditLimit - BLD.DISCOUNT_AMOUNT
   Set BLD = Nothing
   Set m_Rs = Nothing
   
End Function

Private Sub uctlTime1_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTime2_HasChange()
   m_HasModify = True
End Sub
Private Function CheckCredit() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc
Dim Pm As CPayment
Dim TempCreditLimit As Double
Dim CheckMaxCredit As Boolean
Dim TempUserName As String
   
   If Not VerifyDate(lblDocumentDate, uctlDocumentDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If DocumentType = 1 Then
      Dim Cs As CCustomer
      Dim Cus_ID As Long
      Dim SumDB As Double
      Dim TotalSale As Double
      Cus_ID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
      Set Cs = m_Customers(Trim(str(Cus_ID)))
      
      If DocumentType = 1 Then
         If Not (AllowSave) Then
            glbErrorLog.LocalErrorMsg = MapText("ยังไม่มีข้อมูล ใบปะหน้าบัญชี และเวลาได้เกินกำหนดแล้วไม่สามารถสร้างเอกสารใบขายเชื่อได้ " & vbCrLf & "  กรุณาเพิ่มข้อมูลใบปะหน้าบัญชีเพิ่มให้สามารถใช้ Funtion นี้ได้")
            glbErrorLog.ShowUserError
            Exit Function
         End If
      End If
      
      'เพิ่มวงเงินราย Week 01/11/2560 โดยพี่มณ
      If Cs.WEEK_CREDIT_LIMIT > 0 Then
         TotalSale = CheckWeekCreditLimit(uctlDocumentDate.ShowDate)
         If TotalSale > Cs.WEEK_CREDIT_LIMIT Then
            glbErrorLog.LocalErrorMsg = "ยอดหนี้มากกว่าวงเงินรายสัปดาห์ (จำนวน " & FormatNumber(TotalSale - Cs.WEEK_CREDIT_LIMIT) & " ) "
            glbErrorLog.ShowUserError
         Else
            glbErrorLog.LocalErrorMsg = "วงเงินรายสัปดาห์คงเหลือ (จำนวน " & FormatNumber(Cs.WEEK_CREDIT_LIMIT - TotalSale) & " ) "
            glbErrorLog.ShowUserError
         End If
      End If
      'เพิ่มวงเงินราย Week 01/11/2560 โดยพี่มณ
   
      If Cs.CHECK_CREDIT_FLAG = "Y" Then
         CheckMaxCredit = True
         If Cs.SUSPEND_SALES = "Y" Or Cs.CREDIT_LIMIT > 0 Or Cs.MAX_CREDIT > 0 Then
            If Cs.SUSPEND_SALES = "N" Then
               If Cs.CREDIT_LIMIT > 0 Then   'เช็ควงเงิน
                  TempCreditLimit = CheckCreditLimit
                  SumDB = (Cs.CREDIT_LIMIT - TempCreditLimit)
               End If
               
               'เช็ควงวัน
               CheckMaxCredit = CheckMaxCreditDueDate(Cs)
               
            End If
            
            If Cs.SUSPEND_SALES = "Y" Or SumDB < 0 Or (Not CheckMaxCredit) Then
               If Cs.SUSPEND_SALES = "Y" Then
                  glbErrorLog.LocalErrorMsg = "ระงับการขายชั่วคราว"
               ElseIf SumDB < 0 And (Not CheckMaxCredit) Then
                  glbErrorLog.LocalErrorMsg = "ยอดหนี้มากกว่าวงเงิน (จำนวน " & (-SumDB) & " ) และมียอดเกินวงวันสูงสุด"
               ElseIf SumDB < 0 Then
                  glbErrorLog.LocalErrorMsg = "ยอดหนี้มากกว่าวงเงิน (จำนวน " & (-SumDB) & " ) "
               Else
                  glbErrorLog.LocalErrorMsg = "มียอดเกินวงวันสูงสุด"
               End If
            ElseIf SumDB >= 0 Then
               glbErrorLog.LocalErrorMsg = "วงเงินคงเหลือ (จำนวน " & (SumDB) & " ) "
            End If
            
         End If
      ElseIf Cs.CHECK_CREDIT_FLAG = "N" Then
         glbErrorLog.LocalErrorMsg = "ลูกค้ารายนี้ไม่มีการตรวจสอบเครดิต "
      End If
   End If
   glbErrorLog.ShowUserError
End Function

