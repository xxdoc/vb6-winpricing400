VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddEditReturn 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditReturn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000009&
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   1575
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   1635
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboPackageType 
         Height          =   315
         Left            =   9690
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2670
         Width           =   1755
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   2565
         Left            =   150
         TabIndex        =   45
         Top             =   8310
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   4524
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin prjFarmManagement.uctlDate uctlDueDate 
            Height          =   405
            Left            =   7680
            TabIndex        =   48
            Top             =   150
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   714
         End
         Begin prjFarmManagement.uctlTextBox txtPayment 
            Height          =   435
            Left            =   1740
            TabIndex        =   51
            Top             =   1050
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtNote 
            Height          =   435
            Left            =   1740
            TabIndex        =   49
            Top             =   600
            Width           =   1515
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlDate uctlShipDate 
            Height          =   405
            Left            =   7680
            TabIndex        =   52
            Top             =   1050
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   714
         End
         Begin prjFarmManagement.uctlTextBox txtPONo 
            Height          =   435
            Left            =   7680
            TabIndex        =   50
            Top             =   600
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextLookup uctlSellByLookup 
            Height          =   435
            Left            =   1740
            TabIndex        =   53
            Top             =   1530
            Width           =   5385
            _ExtentX        =   9499
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtCredit 
            Height          =   435
            Left            =   1740
            TabIndex        =   47
            Top             =   150
            Width           =   1515
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtTempDONo 
            Height          =   435
            Left            =   1740
            TabIndex        =   55
            Top             =   1980
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   767
         End
         Begin VB.Label lblTempDONo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   60
            TabIndex        =   93
            Top             =   2010
            Width           =   1605
         End
         Begin VB.Label Label11 
            Height          =   315
            Left            =   3360
            TabIndex        =   65
            Top             =   240
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label lblCredit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            TabIndex        =   64
            Top             =   180
            Width           =   1395
         End
         Begin VB.Label lblSellBy 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   30
            TabIndex        =   59
            Top             =   1590
            Width           =   1635
         End
         Begin VB.Label lblPoNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6180
            TabIndex        =   58
            Top             =   630
            Width           =   1395
         End
         Begin VB.Label lblDeliveryPlace 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   210
            TabIndex        =   57
            Top             =   1140
            Width           =   1395
         End
         Begin VB.Label lblShipment 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5640
            TabIndex        =   56
            Top             =   1110
            Width           =   1995
         End
         Begin VB.Label lblNote 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            TabIndex        =   54
            Top             =   630
            Width           =   1395
         End
         Begin VB.Label lblDueDate 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6150
            TabIndex        =   46
            Top             =   180
            Width           =   1395
         End
      End
      Begin VB.ComboBox cboEnpAddress 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2220
         Width           =   9585
      End
      Begin VB.ComboBox cboCustomerAddress 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1770
         Width           =   9585
      End
      Begin VB.ComboBox cboAccount 
         Height          =   315
         Left            =   9090
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   2355
      End
      Begin prjFarmManagement.uctlTextLookup uctlCustomerLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   4
         Top             =   1320
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6570
         TabIndex        =   2
         Top             =   870
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   14
         Top             =   4620
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
         Width           =   2535
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
      Begin GridEX20.GridEX GridEX1 
         Height          =   2565
         Left            =   150
         TabIndex        =   15
         Top             =   5160
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   4524
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
         Column(1)       =   "frmAddEditReturn.frx":27A2
         Column(2)       =   "frmAddEditReturn.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditReturn.frx":290E
         FormatStyle(2)  =   "frmAddEditReturn.frx":2A6A
         FormatStyle(3)  =   "frmAddEditReturn.frx":2B1A
         FormatStyle(4)  =   "frmAddEditReturn.frx":2BCE
         FormatStyle(5)  =   "frmAddEditReturn.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditReturn.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   25
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
         TabIndex        =   10
         Top             =   2640
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtNetTotal 
         Height          =   435
         Left            =   6030
         TabIndex        =   11
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtDiscount 
         Height          =   435
         Left            =   1860
         TabIndex        =   12
         Top             =   3090
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtIncludeDiscount 
         Height          =   435
         Left            =   6030
         TabIndex        =   34
         Top             =   3090
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCashDiscount 
         Height          =   435
         Left            =   1860
         TabIndex        =   39
         Top             =   4590
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtLeft 
         Height          =   435
         Left            =   6030
         TabIndex        =   13
         Top             =   3540
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtCashDiscountAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   61
         Top             =   3540
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   767
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   2565
         Left            =   150
         TabIndex        =   66
         Top             =   8220
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   4524
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.ComboBox cboHold4 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   1740
            Width           =   2355
         End
         Begin VB.ComboBox cboHold3 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   1290
            Width           =   2355
         End
         Begin VB.ComboBox cboHold2 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   840
            Width           =   2355
         End
         Begin VB.ComboBox cboHold1 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   390
            Width           =   2355
         End
         Begin prjFarmManagement.uctlTextBox txtHold1Desc 
            Height          =   435
            Left            =   7080
            TabIndex        =   69
            Top             =   390
            Width           =   4335
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtHold1Amount 
            Height          =   435
            Left            =   4380
            TabIndex        =   68
            Top             =   390
            Width           =   1515
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtHold2Desc 
            Height          =   435
            Left            =   7080
            TabIndex        =   72
            Top             =   840
            Width           =   4335
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtHold2Amount 
            Height          =   435
            Left            =   4380
            TabIndex        =   71
            Top             =   840
            Width           =   1515
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtHold3Desc 
            Height          =   435
            Left            =   7080
            TabIndex        =   75
            Top             =   1290
            Width           =   4335
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtHold3Amount 
            Height          =   435
            Left            =   4380
            TabIndex        =   74
            Top             =   1290
            Width           =   1515
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtHold4Desc 
            Height          =   435
            Left            =   7080
            TabIndex        =   78
            Top             =   1740
            Width           =   4335
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlTextBox txtHold4Amount 
            Height          =   435
            Left            =   4380
            TabIndex        =   77
            Top             =   1740
            Width           =   1515
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin VB.Label lblHold4 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   90
            TabIndex        =   91
            Top             =   1830
            Width           =   1065
         End
         Begin VB.Label lblHold4Desc 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6030
            TabIndex        =   90
            Top             =   1770
            Width           =   945
         End
         Begin VB.Label lblHold4Amount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3570
            TabIndex        =   89
            Top             =   1770
            Width           =   705
         End
         Begin VB.Label lblHold3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   90
            TabIndex        =   88
            Top             =   1380
            Width           =   1065
         End
         Begin VB.Label lblHold3Desc 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6030
            TabIndex        =   87
            Top             =   1320
            Width           =   945
         End
         Begin VB.Label lblHold3Amount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3570
            TabIndex        =   86
            Top             =   1320
            Width           =   705
         End
         Begin VB.Label lblHold2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   90
            TabIndex        =   85
            Top             =   930
            Width           =   1065
         End
         Begin VB.Label lblHold2Desc 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6030
            TabIndex        =   84
            Top             =   870
            Width           =   945
         End
         Begin VB.Label lblHold2Amount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3570
            TabIndex        =   83
            Top             =   870
            Width           =   705
         End
         Begin VB.Label lblHold1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   90
            TabIndex        =   82
            Top             =   480
            Width           =   1065
         End
         Begin VB.Label lblHold1Desc 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6030
            TabIndex        =   81
            Top             =   420
            Width           =   945
         End
         Begin VB.Label lblHold1Amount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3570
            TabIndex        =   80
            Top             =   420
            Width           =   705
         End
         Begin VB.Label Label12 
            Height          =   315
            Left            =   3360
            TabIndex        =   79
            Top             =   300
            Visible         =   0   'False
            Width           =   405
         End
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   2595
         Left            =   150
         TabIndex        =   94
         Top             =   8310
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   4577
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.ComboBox cboPaymentType 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   95
            Top             =   270
            Width           =   2325
         End
         Begin VB.ComboBox cboBank 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   98
            Top             =   720
            Width           =   4035
         End
         Begin VB.ComboBox cboBankBranch 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   101
            Top             =   1170
            Width           =   4035
         End
         Begin prjFarmManagement.uctlTextBox txtCheckNo 
            Height          =   435
            Left            =   7350
            TabIndex        =   96
            Top             =   270
            Width           =   3465
            _ExtentX        =   5001
            _ExtentY        =   767
         End
         Begin prjFarmManagement.uctlDate uctlCheckDate 
            Height          =   405
            Left            =   7350
            TabIndex        =   99
            Top             =   720
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   714
         End
         Begin VB.Label lblBankBranch 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   60
            TabIndex        =   97
            Top             =   1230
            Width           =   1395
         End
         Begin VB.Label lblBank 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   60
            TabIndex        =   104
            Top             =   750
            Width           =   1395
         End
         Begin VB.Label lblPaymentType 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   60
            TabIndex        =   103
            Top             =   330
            Width           =   1395
         End
         Begin VB.Label lblCheckDate 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5820
            TabIndex        =   102
            Top             =   780
            Width           =   1395
         End
         Begin VB.Label lblCheckNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5580
            TabIndex        =   100
            Top             =   330
            Width           =   1665
         End
      End
      Begin prjFarmManagement.uctlTextBox txtDipRcp 
         Height          =   435
         Left            =   6030
         TabIndex        =   105
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
      End
      Begin prjFarmManagement.uctlTextBox txtTotalRcp 
         Height          =   435
         Left            =   1860
         TabIndex        =   106
         Top             =   3960
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   767
      End
      Begin VB.Label Label15 
         Height          =   315
         Left            =   7740
         TabIndex        =   110
         Top             =   4020
         Width           =   585
      End
      Begin VB.Label lblDipRcp 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4290
         TabIndex        =   109
         Top             =   4050
         Width           =   1635
      End
      Begin VB.Label lblTotalRcp 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   108
         Top             =   4080
         Width           =   1425
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   3510
         TabIndex        =   107
         Top             =   4020
         Width           =   765
      End
      Begin VB.Label lblPackageType 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8460
         TabIndex        =   92
         Top             =   2700
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label9 
         Height          =   315
         Left            =   3570
         TabIndex        =   63
         Top             =   3600
         Width           =   765
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   300
         TabIndex        =   62
         Top             =   3630
         Width           =   1425
      End
      Begin VB.Label Label5 
         Height          =   315
         Left            =   3540
         TabIndex        =   60
         Top             =   4650
         Visible         =   0   'False
         Width           =   765
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   6810
         TabIndex        =   19
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
         TabIndex        =   5
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditReturn.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   4410
         TabIndex        =   1
         Top             =   870
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditReturn.frx":3250
         ButtonStyle     =   3
      End
      Begin VB.Label lblCashDiscount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   43
         Top             =   4710
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label10 
         Height          =   315
         Left            =   3540
         TabIndex        =   42
         Top             =   3660
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblLeft 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4710
         TabIndex        =   41
         Top             =   3630
         Width           =   1275
      End
      Begin VB.Label Label8 
         Height          =   315
         Left            =   7680
         TabIndex        =   40
         Top             =   3600
         Width           =   585
      End
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   38
         Top             =   3210
         Width           =   1695
      End
      Begin VB.Label Label6 
         Height          =   315
         Left            =   3540
         TabIndex        =   37
         Top             =   3180
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblIncludeDiscount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4710
         TabIndex        =   36
         Top             =   3180
         Width           =   1275
      End
      Begin VB.Label Label3 
         Height          =   315
         Left            =   7680
         TabIndex        =   35
         Top             =   3150
         Width           =   585
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   10500
         TabIndex        =   3
         Top             =   870
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblEnpAddress 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   2310
         Width           =   1635
      End
      Begin VB.Label lblCustomerAddress 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   1860
         Width           =   1635
      End
      Begin VB.Label lblAccountNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7890
         TabIndex        =   31
         Top             =   1410
         Width           =   1095
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   150
         TabIndex        =   30
         Top             =   1380
         Width           =   1635
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   7710
         TabIndex        =   29
         Top             =   2700
         Width           =   585
      End
      Begin VB.Label lblNetTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5010
         TabIndex        =   28
         Top             =   2730
         Width           =   915
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   3540
         TabIndex        =   27
         Top             =   2730
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5040
         TabIndex        =   26
         Top             =   930
         Width           =   1395
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   20
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditReturn.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   21
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditReturn.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   18
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditReturn.frx":3B9E
         ButtonStyle     =   3
      End
      Begin VB.Label lblTotalAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   24
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   23
         Top             =   930
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditReturn"
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
Private m_Cd As Collection
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
      
      uctlShipDate.ShowDate = m_BillingDoc.SHIPMENT
      txtPayment.Text = m_BillingDoc.PAYMENT_DESC
      txtNote.Text = m_BillingDoc.NOTE
      txtPONo.Text = m_BillingDoc.REF
      chkCommit.Value = FlagToCheck(m_BillingDoc.OLD_COMMIT_FLAG)
      chkCommit.Enabled = (m_BillingDoc.OLD_COMMIT_FLAG = "N")
      Call EnableDisableButton(True)
      
      txtCheckNo.Text = m_BillingDoc.CHECK_NO
      uctlCheckDate.ShowDate = m_BillingDoc.CHECK_DATE
      cboPaymentType.ListIndex = IDToListIndex(cboPaymentType, m_BillingDoc.PAYMENT_TYPE)
      cboBank.ListIndex = IDToListIndex(cboBank, m_BillingDoc.BANK_ID)
      cboBankBranch.ListIndex = IDToListIndex(cboBankBranch, m_BillingDoc.BBRANCH_ID)
      
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
Dim Ri As CReceiptItem
   For Each Di In BD.DoItems
      If Di.Flag = "A" Then
         Di.LINK_ID = GetNextGuiID(BD)
      End If
   Next Di
   
   For Each Ri In BD.ReceiptItems
      If Ri.Flag = "A" Then
         Ri.LINK_ID = GetNextGuiID(BD)
      End If
   Next Ri

End Sub

Private Function GetNextGuiID(BD As CBillingDoc) As Long
Dim Di As CDoItem
Dim Ri As CReceiptItem
Dim MaxId As Long

   MaxId = 0
   For Each Di In BD.DoItems
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   For Each Ri In BD.ReceiptItems
      If Ri.LINK_ID > MaxId Then
         MaxId = Ri.LINK_ID
      End If
   Next Ri
   
   GetNextGuiID = MaxId + 1
End Function

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc
Dim Pm As CPayment
   If ShowMode = SHOW_EDIT Then
      If Area = 1 Then
         If Not VerifyAccessRight("LEDGER_SELL" & "_" & DocumentType & "_" & "EDIT", "แก้ไข") Then
            Call EnableForm(Me, True)
            Exit Function
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
    If Not VerifyDate(lblShipment, uctlShipDate, True) Then
      Exit Function
   End If
   
'   If Not VerifyCombo(lblPaymentType, cboPaymentType, DocumentType <> 2) Then
'      Exit Function
'   End If
   
   If Not CheckUniqueNs(DO_PLAN_UNIQUE, txtDocumentNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      DocAdd = DocAdd + 1
      Call cmdAuto_Click
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
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
   
   m_BillingDoc.CHECK_NO = txtCheckNo.Text
   m_BillingDoc.CHECK_DATE = uctlCheckDate.ShowDate
   m_BillingDoc.PAYMENT_TYPE = cboPaymentType.ItemData(Minus2Zero(cboPaymentType.ListIndex))
   m_BillingDoc.BANK_ID = cboBank.ItemData(Minus2Zero(cboBank.ListIndex))
   If cboBankBranch.ListIndex > 0 Then
      m_BillingDoc.BBRANCH_ID = cboBankBranch.ItemData(Minus2Zero(cboBankBranch.ListIndex))
   End If
      
  m_BillingDoc.SHIPMENT = uctlShipDate.ShowDate
   m_BillingDoc.NOTE = txtNote.Text
   m_BillingDoc.PAYMENT_DESC = txtPayment.Text
   
   Call PopulateGuiID(m_BillingDoc)
   Call glbDaily.GenerateExtraDiscount(m_BillingDoc)
   
   Call EnableForm(Me, False)
   
   If DocumentType = 1 Then
      Call glbDaily.DO2InventoryDoc(m_BillingDoc, Ivd, Area, 10)
   ElseIf DocumentType = 2 Then  'ใบเสร็จ
      Call glbDaily.DO2InventoryDoc(m_BillingDoc, Ivd, Area, 21)
      Call glbDaily.DO2Payment(m_BillingDoc, Pm)
   ElseIf DocumentType = 18 Then  'ใบรับคืนสินค้า
      Call glbDaily.RT2InventoryDoc(m_BillingDoc, Ivd, Area, 31)
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
            If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo, False) Then
               Exit Sub
            End If
            frmAddReturnItem.CusID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
         End If

         Set frmAddReturnItem.TempCollection = m_BillingDoc.ReceiptItems
         frmAddReturnItem.ShowMode = SHOW_ADD
         frmAddReturnItem.HeaderText = MapText("เพิ่มรายการใบรับคืนสินค้า")
        
         Load frmAddReturnItem
         frmAddReturnItem.Show 1
   
         OKClick = frmAddReturnItem.OKClick
   
         Unload frmAddReturnItem
         Set frmAddReturnItem = Nothing
   
         If OKClick Then
            Call GetTotalPrice
   
            GridEX1.ItemCount = CountItem(m_BillingDoc.ReceiptItems)
            GridEX1.Rebind
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
Private Function GetDocumentNo(DocNoType As Long) As String
Dim DOC_ID As Long
Dim Cd As CConfigDoc
Dim TempStr As String
Dim I As Long
Dim ServerDateTime As String

   DOC_ID = SELL_RETURN
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
'         GetDocumentNo = GetDocumentNo & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
'         m_BillingDoc.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
'         m_BillingDoc.CONFIG_DOC_TYPE = SELL_RETURN
         
          If Cd.GetFieldValue("AUTO_BEGIN_FLAG") = "Y" Then
               If CheckNewMounth And CheckUniqueNs(DO_PLAN_UNIQUE, GetDocumentNo & Format(1, TempStr), ID) Then
                  GetDocumentNo = GetDocumentNo & Format(1, TempStr) 'เริ่มจาก 1 เสมอ
                  m_BillingDoc.RUNNING_NO = 1
               Else
                  GetDocumentNo = GetDocumentNo & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
                 m_BillingDoc.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
               End If
          Else
               GetDocumentNo = GetDocumentNo & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
                m_BillingDoc.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
          End If
          m_BillingDoc.CONFIG_DOC_TYPE = DOC_ID
      Else
         GetDocumentNo = ""
      End If
   End If
End Function
Private Sub cmdAuto_Click()
   If Trim(txtDocumentNo.Text) = "" And ShowMode = SHOW_ADD Then
      txtDocumentNo.Text = GetDocumentNo(DocumentType)
   End If
'Dim ID As Long
'Dim Cd As CConfigDoc
'Dim TempStr As String
'Dim I As Long
'Dim ServerDateTime As String
'
'   ID = SELL_RETURN
'   If ID > 0 Then
'      Set Cd = GetObject("CConfigDoc", m_Cd, Trim(str(ID)), False)
'      If Not (Cd Is Nothing) Then
'         txtDocumentNo.Text = Cd.GetFieldValue("PREFIX") & Cd.GetFieldValue("CODE1")
'         TempStr = ""
'         If Cd.GetFieldValue("YEAR_TYPE") = 1 Then
'            TempStr = Right(Format(Year(Now) + 543, "0000"), 2)
'         ElseIf Cd.GetFieldValue("YEAR_TYPE") = 2 Then
'            TempStr = Format(Year(Now) + 543, "0000")
'         ElseIf Cd.GetFieldValue("YEAR_TYPE") = 3 Then
'            TempStr = Right(Format(Year(Now), "0000"), 2)
'         ElseIf Cd.GetFieldValue("YEAR_TYPE") = 4 Then
'            TempStr = Format(Year(Now), "0000")
'         End If
'         txtDocumentNo.Text = txtDocumentNo.Text & TempStr & Cd.GetFieldValue("CODE2")
'         TempStr = ""
'         If Cd.GetFieldValue("MONTH_TYPE") = 1 Then
'            TempStr = Format(Month(Now), "00")
'         End If
'         txtDocumentNo.Text = txtDocumentNo.Text & TempStr & Cd.GetFieldValue("CODE3")
'         TempStr = ""
'         For I = 1 To Cd.GetFieldValue("DIGIT_AMOUNT")
'            TempStr = TempStr & "0"
'         Next I
'         txtDocumentNo.Text = txtDocumentNo.Text & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
'         m_BillingDoc.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
'         m_BillingDoc.CONFIG_DOC_TYPE = SELL_RETURN
'      Else
'         txtDocumentNo.Text = ""
'      End If
'   End If
End Sub

Private Sub cmdCustomer_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim OKClick As Boolean
Dim TempCol As Collection
Dim Cs As CCustomer

   Set TempCol = New Collection
   
   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("ค้นหา", "-", "เพิ่มข้อมูลใหม่")
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
         m_BillingDoc.ReceiptItems.Remove (ID2)
      Else
         m_BillingDoc.ReceiptItems.Item(ID2).Flag = "D"
      End If

      Call GetTotalPrice
      GridEX1.ItemCount = CountItem(m_BillingDoc.ReceiptItems)
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
      frmAddEditDoItem.DocumentType = DocumentType
      frmAddEditDoItem.Area = Area
      frmAddEditDoItem.ID = ID
      frmAddEditDoItem.COMMIT_FLAG = m_BillingDoc.OLD_COMMIT_FLAG
      Set frmAddEditDoItem.TempCollection = m_BillingDoc.ReceiptItems
      frmAddEditDoItem.HeaderText = MapText("แก้ไขรายการรับคืนสินค้า")
      
      frmAddEditDoItem.ParentShowMode = ShowMode
      frmAddEditDoItem.ShowMode = SHOW_EDIT
      Load frmAddEditDoItem
      frmAddEditDoItem.Show 1

      OKClick = frmAddEditDoItem.OKClick

      Unload frmAddEditDoItem
      Set frmAddEditDoItem = Nothing

      If OKClick Then
         Call GetTotalPrice
         GridEX1.ItemCount = CountItem(m_BillingDoc.ReceiptItems)
         GridEX1.Rebind
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

Private Sub CalculateIncludePrice()
Dim II As CLotItem
Dim AvgFee As Double

'   If m_SumUnit > 0 Then
'      AvgFee = Val(txtTotalAmount.Text) / m_SumUnit
'   Else
'      AvgFee = 0
'   End If
'
'   For Each II In m_BillingDoc.DoItems
'      If II.Flag <> "D" Then
'         II.INCLUDE_UNIT_PRICE = II.ACTUAL_UNIT_PRICE + AvgFee
'      End If
'   Next II
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
   End If

   If Not Report Is Nothing Then
      Call Report.AddParam(lMenuChosen, "REPORT_TYPE")
      Call Report.AddParam(m_BillingDoc.BILLING_DOC_ID, "BILLING_DOC_ID")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
      Call Report.AddParam(MapText("ใบส่งสินค้า/ใบแจ้งหนี้"), "REPORT_HEADER")
      Call Report.AddParam(DocumentType, "DOCUMENT_TYPE")
      
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
      
      Call EnableForm(Me, False)
      Call LoadEnterpriseAddress(cboEnpAddress, , , True)
      
      Call InitPaymentType(cboPaymentType)
      Call LoadBank(cboBank)
      
      Call LoadConfigDoc(Nothing, m_Cd)
      
      Call InitPackageType(cboPackageType)
      
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
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่ใบแจ้งหนี้"))
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
'   Call InitNormalLabel(lblTotalDiscount, MapText("น้ำหนักรวม"))
   Call InitNormalLabel(Label1, MapText("ตัว"))
'   Call InitNormalLabel(Label2, MapText("ก.ก."))
   Call InitNormalLabel(Label4, MapText("บาท"))
   Call InitNormalLabel(Label5, MapText("%"))
   Call InitNormalLabel(lblNetTotal, MapText("ราคารวม"))
   Call InitNormalLabel(lblDiscount, MapText("ส่วนลด"))
   Call InitNormalLabel(lblCashDiscount, MapText("% ส่วนลดเงินสด"))
   Call InitNormalLabel(lblIncludeDiscount, MapText("รวมส่วนลด"))
   Call InitNormalLabel(lblLeft, MapText("คงค้าง"))
   Call InitCheckBox(chkCommit, "คำนวณ")
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
   
   
   txtDocumentNo.Enabled = False
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
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdCustomer, MapText("F"))
   
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
      If m_BillingDoc.ReceiptItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CReceiptItem
      If m_BillingDoc.ReceiptItems.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_BillingDoc.ReceiptItems, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.RECEIPT_ITEM_ID
      Values(2) = RealIndex
      Values(3) = CR.ShowDescText
      Values(4) = FormatNumber(CR.RETURN_AMOUNT)
      Values(5) = FormatNumber(CR.AVG_PRICE)
      Values(6) = FormatNumber(CR.RETURN_TOTAL_PRICE)
      Values(7) = CR.DOCUMENT_NO
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
      GridEX1.ItemCount = CountItem(m_BillingDoc.ReceiptItems)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If DocumentType = 18 Then
         Call EnableDisableButton(False)
      Else
         Call EnableDisableButton(True)
      End If
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

   CustomerID = uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex))
   If CustomerID > 0 Then
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
